import re
import shutil
from pathlib import Path
from pptx2md import convert, ConversionConfig

class SlidevConverter:
    def __init__(self):
        # Patterns for cleaning up content
        self.bold_pattern = re.compile(r'__([^_]+)__')
        self.italic_pattern = re.compile(r'_([^_]+)_')
        # Match both empty alt text and any alt text in images
        self.image_pattern = re.compile(r'!\[([^\]]*)\]\(([^)]+)\)')
        self.weird_chars_pattern = re.compile(r'â€™|â€œ|â€|â€˜')
        self.bullet_pattern = re.compile(r'^\s*\*\s*__([^_]+)__(.*)$', re.MULTILINE)

    def clean_text(self, text):
        """Clean up weird characters and formatting"""
        # Fix smart quotes and other weird characters
        text = text.replace('â€™', "'")
        text = text.replace('â€œ', '"')
        text = text.replace('â€', '"')
        text = text.replace('â€˜', "'")
        text = text.replace('\\,', ',')
        text = text.replace('\\', '')  # Remove stray backslashes

        # Convert double underscores to single underscores for italics
        # But be careful not to affect headers or bold text patterns
        text = self.convert_italics(text)

        return text

    def convert_italics(self, text):
        """Convert double underscore italics to single underscore italics"""
        # Pattern to match _text_ that should become *text* (avoiding headers and bold)
        # Look for single underscores around words that aren't at start of line
        single_underscore_pattern = re.compile(r'(?<!_)_([^_\s][^_]*[^_\s])_(?!_)')

        # Convert single underscores to asterisks for italics
        text = single_underscore_pattern.sub(r'*\1*', text)

        return text

    def extract_title_from_content(self, content):
        """Extract the main title from slide content"""
        lines = content.split('\n')
        for line in lines:
            line = line.strip()
            if line and not line.startswith('#') and not line.startswith('*') and not line.startswith('!'):
                # Clean the line of markdown formatting
                title = re.sub(r'[*_]', '', line)
                title = self.clean_text(title)
                return title
        return "Slide"

    def detect_slide_type(self, content):
        """Detect what type of slide this should be"""
        lines = [line.strip() for line in content.split('\n') if line.strip()]

        # Check if it's a title/cover slide
        if len(lines) <= 3 and any('Database Systems' in line or 'CSCI' in line for line in lines):
            return 'cover'

        # Check if it has images (should be two-cols layout)
        if self.image_pattern.search(content):
            return 'two-cols'

        # Check if content is very long (should be full layout)
        if len(content) > 800:
            return 'full'

        return 'default'

    def convert_image_path(self, image_line, presentation_name, slide_number=1):
        """Convert pptx2md image paths to Slidev-compatible paths with URL encoding"""
        img_match = self.image_pattern.search(image_line)
        if img_match:
            alt_text = img_match.group(1)
            original_path = img_match.group(2)

            # Convert Windows path separators but keep URL encoding for spaces
            # %5C -> / but keep %20 for spaces
            clean_path = original_path.replace('%5C', '/')
            filename = clean_path.split('/')[-1]

            # Create path that matches the working format: ./img/encoded_name/encoded_filename
            if filename:
                # URL encode the presentation name for consistency
                encoded_presentation_name = presentation_name.replace(' ', '%20')
                simple_path = f"./img/{encoded_presentation_name}/{filename}"
            else:
                encoded_presentation_name = presentation_name.replace(' ', '%20')
                simple_path = f"./img/{encoded_presentation_name}/slide_{slide_number}.png"

            return f"![{alt_text}]({simple_path})"
        return image_line

    def process_content_for_slide(self, content, presentation_name, slide_number=1):
        """Process content for a slide, handling two-column layout"""
        content = self.clean_text(content)

        # Split content into main content and images
        lines = content.split('\n')
        main_content = []
        images = []

        i = 0
        while i < len(lines):
            line = lines[i].strip()

            if not line:  # Skip empty lines initially, add back later if needed
                i += 1
                continue

            # Handle images - convert to simple Slidev format with presentation folder
            img_match = self.image_pattern.search(line)
            if img_match:
                converted_image = self.convert_image_path(line, presentation_name, slide_number)
                images.append(converted_image)
                i += 1
                continue

            # Convert bold headers to actual headers
            if line.startswith('__') and line.endswith('__') and len(line) > 4:
                # This is likely a header
                header_text = line.strip('_')
                header_text = self.clean_text(header_text)
                main_content.append(f"# {header_text}")
            elif line.startswith('*'):
                # Handle bullet points
                if '__' in line:
                    # Bullet point with bold text
                    bullet_match = re.match(r'^\s*\*\s*__([^_]+)__(.*)$', line)
                    if bullet_match:
                        bold_text = bullet_match.group(1)
                        rest_text = bullet_match.group(2).strip()
                        if rest_text:
                            main_content.append(f"* **{bold_text}**{rest_text}")
                        else:
                            main_content.append(f"* **{bold_text}**")
                    else:
                        main_content.append(line)
                else:
                    # Regular bullet point
                    main_content.append(line)
            else:
                # Regular content - check if it should be a header
                if line and not line.startswith('#'):
                    # If line is all caps or looks like a title, make it a header
                    cleaned_line = re.sub(r'[*_]', '', line).strip()
                    if len(cleaned_line) < 100 and (cleaned_line.isupper() or
                                                    any(keyword in cleaned_line.lower() for keyword in ['history', 'what is', 'why', 'how'])):
                        main_content.append(f"# {cleaned_line}")
                    else:
                        main_content.append(line)
                else:
                    main_content.append(line)

            i += 1

        # Clean up main content - remove excessive empty lines
        cleaned_main = []
        prev_empty = False
        for line in main_content:
            if line.strip():
                cleaned_main.append(line)
                prev_empty = False
            elif not prev_empty:
                cleaned_main.append('')
                prev_empty = True

        # If we have images, create two-column layout
        if images:
            result = '\n'.join(cleaned_main)
            result += '\n\n::right::\n\n'
            result += '\n\n'.join(images)
            return result, 'two-cols'
        else:
            return '\n'.join(cleaned_main), 'default'

    def create_slidev_header(self, title):
        """Create the Slidev header with proper frontmatter"""
        header = f"""---
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
        return header

    def convert_to_slidev(self, markdown_content, title, presentation_name):
        """Convert extracted markdown to Slidev format"""
        # Split content into slides - but be careful about splitting
        raw_slides = markdown_content.split('---')

        # Start with Slidev header
        slidev_content = self.create_slidev_header(title)

        processed_slides = 0
        for i, slide_content in enumerate(raw_slides):
            slide_content = slide_content.strip()
            if not slide_content:
                continue

            # Skip if this looks like metadata or very short content
            if len(slide_content) < 10:
                continue

            # Process the slide content with presentation name and slide number for image naming
            processed_content, layout = self.process_content_for_slide(slide_content, presentation_name, processed_slides + 1)

            if not processed_content.strip():
                continue

            processed_slides += 1

            # Add slide break with layout if needed
            if layout == 'full':
                slidev_content += "---\nlayout: full\n---\n\n"
            elif layout == 'two-cols':
                slidev_content += "---\n\n"  # Use default two-cols layout
            else:
                slidev_content += "---\n\n"

            slidev_content += processed_content + "\n\n"

        # Add end slide only if we processed some slides
        if processed_slides > 0:
            slidev_content += "---\nlayout: end\n---\n"

        return slidev_content

def convert_presentations():
    """Main conversion function"""
    converter = SlidevConverter()
    presentations_dir = Path('./presentations')
    output_dir = Path('presentation-conversion')

    # Check if directories exist
    if not presentations_dir.exists():
        print(f"Error: {presentations_dir} directory does not exist")
        return

    output_dir.mkdir(exist_ok=True)

    pptx_files = list(presentations_dir.glob('*.pptx'))
    if not pptx_files:
        print("No .pptx files found in presentations directory")
        return

    for pptx_file in pptx_files:
        print(f"Processing {pptx_file.name}")
        basename = pptx_file.stem
        md_path = output_dir / f"{basename}.md"
        img_dir = output_dir / "img" / basename
        img_dir.mkdir(parents=True, exist_ok=True)

        try:
            # Convert PPTX to markdown using original structure
            convert(
                ConversionConfig(
                    pptx_path=pptx_file,
                    output_path=md_path,
                    image_dir=img_dir,
                    disable_notes=True,
                    enable_slides=True
                )
            )

            # Verify the markdown file was created
            if not md_path.exists():
                print(f"Error: Markdown file was not created for {pptx_file.name}")
                continue

            # Read the extracted markdown
            try:
                with open(md_path, 'r', encoding='utf-8') as f:
                    markdown_content = f.read()
            except UnicodeDecodeError:
                # Try with different encoding
                with open(md_path, 'r', encoding='latin1') as f:
                    markdown_content = f.read()

            if not markdown_content.strip():
                print(f"Warning: Empty content extracted from {pptx_file.name}")
                continue

            # Convert to Slidev format
            title = basename.replace('_', ' ').replace('-', ' ').title()
            slidev_content = converter.convert_to_slidev(markdown_content, title, basename)

            # Write Slidev markdown back to the same file
            with open(md_path, 'w', encoding='utf-8') as f:
                f.write(slidev_content)

            print(f"Successfully converted {pptx_file.name} to Slidev format")

        except ImportError as e:
            print(f"Error: Missing required library. {e}")
            print("Please install pptx2md: pip install pptx2md")
            break
        except Exception as e:
            print(f"Failed to convert {pptx_file.name}: {e}")
            import traceback
            traceback.print_exc()

if __name__ == "__main__":
    convert_presentations()
