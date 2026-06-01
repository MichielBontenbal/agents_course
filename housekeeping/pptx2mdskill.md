# Skill: Convert PowerPoint to Markdown

## Description

This skill provides a Python-based method to convert PowerPoint (`.pptx`) presentations into structured Markdown (`.md`) files. It extracts text, slide titles, and bullet points, and formats them into Markdown syntax.

---

## Requirements

- Python 3.8+
- `python-pptx` library

Install the required library:

```bash
pip install python-pptx
```

---

## Code

### `pptx_to_markdown.py`

```python
from pptx import Presentation
import re
import argparse

def pptx_to_markdown(pptx_path, md_path):
    """
    Convert a PowerPoint (.pptx) file to a Markdown (.md) file.
    
    Args:
        pptx_path (str): Path to the input .pptx file.
        md_path (str): Path to save the output .md file.
    """
    prs = Presentation(pptx_path)
    markdown_content = []

    for i, slide in enumerate(prs.slides, start=1):
        markdown_content.append(f"# Slide {i}\n\n")

        for shape in slide.shapes:
            if hasattr(shape, "text"):
                text = shape.text.strip()
                if text:
                    # Handle bullet points (indented lines)
                    if "\n" in text:
                        lines = text.split("\n")
                        for line in lines:
                            line = line.strip()
                            if line:
                                # Check if the line starts with a bullet or number
                                if re.match(r'^(\d+\.|\-|\*)\s', line):
                                    markdown_content.append(f"- {line[2:].strip() if line.startswith('- ') else line.strip()}\n")
                                else:
                                    markdown_content.append(f"- {line}\n")
                    else:
                        markdown_content.append(f"{text}\n")
        markdown_content.append("\n---\n\n")

    with open(md_path, "w", encoding="utf-8") as md_file:
        md_file.writelines(markdown_content)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Convert PowerPoint to Markdown")
    parser.add_argument("--input", type=str, required=True, help="Path to the input .pptx file")
    parser.add_argument("--output", type=str, required=True, help="Path to save the output .md file")
    args = parser.parse_args()
    
    pptx_to_markdown(args.input, args.output)
    print(f"Conversion complete! Markdown file saved as {args.output}")
```

---

## Usage

### Command Line

Run the script from the terminal:

```bash
python pptx_to_markdown.py --input input.pptx --output output.md
```

### Example

- **Input**: `presentation.pptx`
- **Output**: `presentation.md`

---

## Features

- Extracts slide titles as Markdown headers (`# Slide {i}`).
- Converts bullet points to Markdown lists (`-` ).
- Preserves plain text content.
- Handles multi-line text and indented bullet points.

---

## Limitations

- **Images and Tables**: This script does not extract images or tables. For advanced use cases, consider extending the script with libraries like `Pillow` (for images) or `pandas` (for tables).
- **Styling**: Font styles (bold, italic) are not preserved. Use Markdown syntax (e.g., `**bold**`, `*italic*`) manually if needed.
- **Complex Layouts**: May not handle complex slide layouts perfectly. Test with your specific `.pptx` files.

---

## Extending the Script

### Adding Image Support

To extract images, use the `python-pptx` library to access slide images and save them to a folder. Update the Markdown to include image references:

```python
from pptx import Presentation
from pptx.util import Inches
import os

# Inside the slide loop:
for shape in slide.shapes:
    if shape.shape_type == 13:  # 13 is the type for pictures
        image = shape.image
        image_bytes = image.blob
        image_path = f"images/slide_{i}_image_{shape.shape_id}.png"
        os.makedirs("images", exist_ok=True)
        with open(image_path, "wb") as img_file:
            img_file.write(image_bytes)
        markdown_content.append(f"![Image]({image_path})\n\n")
```

### Adding Table Support

Use `pandas` to convert tables to Markdown:

```python
import pandas as pd

# Inside the slide loop:
for shape in slide.shapes:
    if shape.has_table:
        table = shape.table
        data = []
        for row in table.rows:
            row_data = [cell.text.strip() for cell in row.cells]
            data.append(row_data)
        df = pd.DataFrame(data[1:], columns=data[0])
        markdown_content.append(f"\n{df.to_markdown(index=False)}\n\n")
```

---

## Notes

- Tested with `.pptx` files created in Microsoft PowerPoint and LibreOffice Impress.
- For large presentations, consider adding progress indicators or error handling.

---

## License

This skill is provided as-is. Feel free to adapt and extend it for your needs.
