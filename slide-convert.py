import re
from pathlib import Path
from pptx2md import convert, ConversionConfig


class SlidevConverter:
    def __init__(self):
        self.image_pattern = re.compile(r'!\[([^\]]*)\]\(([^)]+)\)')

    def clean_text(self, text):
        """Clean up weird characters and formatting artifacts"""
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

        # Remove stray backslashes
        text = re.sub(r'\\(?![\\*_\[\]])', '', text)

        # Collapse multiple spaces
        text = re.sub(r'  +', ' ', text)

        return text

    def normalize_formatting(self, text):
        """Normalize bold and italic formatting to consistent syntax"""
        # Remove problematic patterns like __  ** text **  __
        text = re.sub(r'__\s*\*\*\s*([^*]+?)\s*\*\*\s*__', r'**\1**', text)
        text = re.sub(r'__\s*_([^_]+?)_\s*__', r'*\1*', text)

        # Fix malformed patterns like *l*oad*__ or _**evealed*__
        text = re.sub(r'\*([a-z])\*([a-z]+)\*__', r'**\1\2**', text)
        text = re.sub(r'_\*\*([a-z]+)\*__', r'**\1**', text)
        text = re.sub(r'\*_\*\*([^*]+?)\*__', r'**\1**', text)

        # Convert standalone __ to ** for bold
        text = re.sub(r'__([^_\s][^_]*?[^_\s])__', r'**\1**', text)

        # Convert single _ to * for italics
        text = re.sub(r'(?<!\w)_([^_\s][^_]*?[^_\s])_(?!\w)', r'*\1*', text)

        # Clean up spacing around formatting
        text = re.sub(r'\*\*\s+', '**', text)
        text = re.sub(r'\s+\*\*', '**', text)
        text = re.sub(r'\*\s+', '*', text)
        text = re.sub(r'\s+\*(?!\*)', '*', text)

        # Remove empty formatting
        text = re.sub(r'\*\*\s*\*\*', '', text)
        text = re.sub(r'\*\s*\*', '', text)

        return text

    def is_likely_header(self, line):
        """Determine if a line should be a header"""
        line = line.strip()

        if line.startswith('#'):
            return True

        # Remove formatting to check content
        clean = re.sub(r'[*_]', '', line).strip()

        if len(clean) < 3 or len(clean) > 120:
            return False

        # Check for header indicators
        is_short_statement = len(clean) < 80 and clean.count('.') == 0
        has_header_keywords = any(kw in clean.lower() for kw in [
            'background job', 'the web side', 'what is', 'why', 'how',
            'choosing', 'part ', 'lesson', 'introduction', 'history'
        ])
        ends_with_colon = clean.endswith(':')
        is_question = clean.endswith('?')

        return (is_short_statement and (has_header_keywords or ends_with_colon)) or is_question

    def process_line(self, line):
        """Process a single line appropriately"""
        line = line.rstrip()

        if not line.strip():
            return ''

        # Already a header
        if line.strip().startswith('#'):
            clean = self.clean_text(line)
            return self.normalize_formatting(clean)

        # Check for bullet point FIRST
        bullet_match = re.match(r'^(\s*)\*\s+(.+)$', line)
        if bullet_match:
            content = bullet_match.group(2)

            # Clean and normalize the content
            content = self.clean_text(content)
            content = self.normalize_formatting(content)

            # Ensure proper indentation (2 spaces for bullets)
            return f"  * {content}"

        # Check if should be header
        if self.is_likely_header(line):
            clean = re.sub(r'^[*\s]*', '', line)  # Remove leading * and spaces
            clean = re.sub(r'[*_]+([^*_]+)[*_]+', r'\1', clean)  # Remove formatting
            clean = self.clean_text(clean)
            return f"# {clean}"

        # Regular paragraph text - convert to bullet point by default
        processed = self.clean_text(line)
        processed = self.normalize_formatting(processed)

        # Don't wrap entire paragraphs in bold
        if processed.startswith('**') and processed.endswith('**') and processed.count('**') == 2:
            processed = processed[2:-2]

        # Convert regular text lines to bullet points
        return f"  * {processed}"

    def convert_image_path(self, image_line, presentation_name, slide_number=1):
        """Convert pptx2md image paths to Slidev-compatible paths"""
        img_match = self.image_pattern.search(image_line)
        if img_match:
            alt_text = img_match.group(1)
            original_path = img_match.group(2)

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
        lines = content.split('\n')
        main_content = []
        images = []

        for line in lines:
            # Skip empty lines during processing
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

        # Clean up consecutive duplicate bullets and excessive empty lines
        cleaned = []
        prev_line = None
        prev_empty = False

        for line in main_content:
            # Skip if exact duplicate of previous line
            if line == prev_line:
                continue

            if line.strip():
                cleaned.append(line)
                prev_empty = False
                prev_line = line
            elif not prev_empty:
                cleaned.append('')
                prev_empty = True
                prev_line = None

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