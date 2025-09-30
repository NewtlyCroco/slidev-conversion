import re
import shutil
from pathlib import Path
from pptx2md import convert, ConversionConfig

class SlidevConverter:
    def __init__(self):
        # Patterns for cleaning up content
        self.image_pattern = re.compile(r'!\[([^\]]*)\]\(([^)]+)\)')
        # Match bold patterns that shouldn't be headers
        self.inline_bold_pattern = re.compile(r'\*\*([^*\n]+)\*\*')

    def clean_text(self, text):
        """Clean up weird characters and formatting"""
        # Fix smart quotes and other weird characters
        replacements = {
            'Ã¢â‚¬â„¢': "'",
            'Ã¢â‚¬Å"': '"',
            'Ã¢â‚¬': '"',
            'Ã¢â‚¬Ëœ': "'",
            'â€œ': '"',
            'â€': '"',
            'â€™': "'",
            '\\,': ',',
        }

        for old, new in replacements.items():
            text = text.replace(old, new)

        # Remove stray backslashes but be careful with intentional escapes
        text = re.sub(r'\\(?![\\*_\[\]])', '', text)

        return text

    def is_likely_header(self, line):
        """Determine if a line should be a header"""
        line = line.strip()

        # Already a header
        if line.startswith('#'):
            return True

        # Remove formatting to check content
        clean = re.sub(r'[*_]', '', line).strip()

        # Empty or very short
        if len(clean) < 3:
            return False

        # Too long to be a header
        if len(clean) > 100:
            return False

        # Check for header-like qualities
        header_indicators = [
            clean.isupper(),  # All caps
            clean.endswith(('?', ':')),  # Questions or labels
            any(keyword in clean.lower() for keyword in [
                'history', 'what is', 'why', 'how', 'choosing',
                'part ', 'lesson', 'introduction', 'conclusion'
            ]),
            # Single sentence without much punctuation
            clean.count('.') <= 1 and clean.count(',') <= 2
        ]

        return any(header_indicators)

    def process_line(self, line):
        """Process a single line, converting it appropriately"""
        line = line.rstrip()

        # Empty line
        if not line.strip():
            return ''

        # Already formatted as header
        if line.strip().startswith('#'):
            return self.clean_text(line)

        # Check if this should be a header
        if self.is_likely_header(line):
            # Remove bold/italic formatting
            clean = re.sub(r'\*+([^*]+)\*+', r'\1', line.strip())
            clean = re.sub(r'__([^_]+)__', r'\1', clean)
            return f"# {self.clean_text(clean)}"

        # Bullet point
        if line.strip().startswith('*'):
            # Proper indentation (2 spaces)
            indent = len(line) - len(line.lstrip())
            bullet_match = re.match(r'^(\s*)\*\s*(.+)$', line)
            if bullet_match:
                content = bullet_match.group(2)
                # Convert __text__ to **text** for bold in bullets
                content = re.sub(r'__([^_]+)__', r'**\1**', content)
                # Clean up excessive bold
                content = re.sub(r'\*\*\s*\*\*', '', content)
                return f"  * {self.clean_text(content)}"

        # Regular paragraph text - remove excessive bold
        # Only keep bold for actual emphasis, not whole sentences
        processed = re.sub(r'^\*\*(.+)\*\*$', r'\1', line.strip())
        processed = re.sub(r'__(.+)__', r'*\1*', processed)  # Convert __ to * for italics

        return self.clean_text(processed)

    def convert_image_path(self, image_line, presentation_name, slide_number=1):
        """Convert pptx2md image paths to Slidev-compatible paths"""
        img_match = self.image_pattern.search(image_line)
        if img_match:
            alt_text = img_match.group(1)
            original_path = img_match.group(2)

            # Convert path separators
            clean_path = original_path.replace('%5C', '/')
            filename = clean_path.split('/')[-1]

            if filename:
                encoded_name = presentation_name.replace(' ', '%20')
                simple_path = f"./img/{encoded_name}/{filename}"
            else:
                encoded_name = presentation_name.replace(' ', '%20')
                simple_path = f"./img/{encoded_name}/slide_{slide_number}.png"

            return f"![{alt_text}]({simple_path})"
        return image_line

    def process_slide_content(self, content, presentation_name, slide_number=1):
        """Process content for a slide"""
        content = self.clean_text(content)

        lines = content.split('\n')
        main_content = []
        images = []

        for line in lines:
            # Skip empty lines initially
            if not line.strip():
                continue

            # Handle images
            if self.image_pattern.search(line):
                converted = self.convert_image_path(line, presentation_name, slide_number)
                images.append(converted)
                continue

            # Process the line
            processed = self.process_line(line)
            if processed:
                main_content.append(processed)

        # Remove consecutive empty lines
        cleaned = []
        prev_empty = False
        for line in main_content:
            if line.strip():
                cleaned.append(line)
                prev_empty = False
            elif not prev_empty:
                cleaned.append('')
                prev_empty = True

        # Build result
        if images:
            result = '\n'.join(cleaned)
            result += '\n\n::right::\n\n'
            result += '\n\n'.join(images)
            return result, 'two-cols'
        else:
            return '\n'.join(cleaned), 'default'

    def create_slidev_header(self, title):
        """Create the Slidev header"""
        return f"""---
defaults:
  layout: two-cols
mdc: true
fonts:
  mono: Cascadia Mono
  sans: Atkinson Hyperlegible
layout: cover
---

# {title}
"""

    def convert_to_slidev(self, markdown_content, title, presentation_name):
        """Convert to Slidev format"""
        raw_slides = markdown_content.split('---')
        slidev_content = self.create_slidev_header(title)

        processed_slides = 0
        for i, slide_content in enumerate(raw_slides):
            slide_content = slide_content.strip()
            if not slide_content or len(slide_content) < 10:
                continue

            processed_content, layout = self.process_slide_content(
                slide_content, presentation_name, processed_slides + 1
            )

            if not processed_content.strip():
                continue

            processed_slides += 1

            # Add slide separator
            if layout == 'full':
                slidev_content += "---\nlayout: full\n---\n\n"
            else:
                slidev_content += "---\n\n"

            slidev_content += processed_content + "\n"

        if processed_slides > 0:
            slidev_content += "\n---\nlayout: end\n---\n"

        return slidev_content

def convert_presentations():
    """Main conversion function"""
    converter = SlidevConverter()
    presentations_dir = Path('./presentations')
    output_dir = Path('presentation-conversion')

    if not presentations_dir.exists():
        print(f"Error: {presentations_dir} directory does not exist")
        return

    output_dir.mkdir(exist_ok=True)

    pptx_files = list(presentations_dir.glob('*.pptx'))
    if not pptx_files:
        print("No .pptx files found")
        return

    for pptx_file in pptx_files:
        print(f"Processing {pptx_file.name}")
        basename = pptx_file.stem
        md_path = output_dir / f"{basename}.md"
        img_dir = output_dir / "img" / basename
        img_dir.mkdir(parents=True, exist_ok=True)

        try:
            convert(
                ConversionConfig(
                    pptx_path=pptx_file,
                    output_path=md_path,
                    image_dir=img_dir,
                    disable_notes=True,
                    enable_slides=True
                )
            )

            if not md_path.exists():
                print(f"Error: Markdown file not created")
                continue

            try:
                with open(md_path, 'r', encoding='utf-8') as f:
                    markdown_content = f.read()
            except UnicodeDecodeError:
                with open(md_path, 'r', encoding='latin1') as f:
                    markdown_content = f.read()

            if not markdown_content.strip():
                print(f"Warning: Empty content")
                continue

            title = basename.replace('_', ' ').replace('-', ' ').title()
            slidev_content = converter.convert_to_slidev(
                markdown_content, title, basename
            )

            with open(md_path, 'w', encoding='utf-8') as f:
                f.write(slidev_content)

            print(f"Successfully converted {pptx_file.name}")

        except Exception as e:
            print(f"Failed to convert {pptx_file.name}: {e}")
            import traceback
            traceback.print_exc()

if __name__ == "__main__":
    convert_presentations()
