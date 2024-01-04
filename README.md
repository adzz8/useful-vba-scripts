# vba scripts

This script allows users to automatically convert multiple PowerPoint files (`.ppt` and `.pptx`) to PDF format, preserving embedded fonts. It's particularly useful for cases where there's a need to distribute presentations without losing custom font appearances.

This script converts all fonts used in a PowerPoint file to a new target font and saves the files. It's useful for cases where you want to switch the presentation from a font that is read only embedded where you need it to be read and write embedded.

## How to Use

1. Open one of the PowerPoint presentations.
2. Press `ALT + F11` to access the VBA editor in PowerPoint.
3. Insert a new module by right-clicking on "VBAProject" > Insert > Module.
4. Copy and paste the provided VBA script into this module.
5. Close the VBA editor.
6. Press `ALT + F8`, select "[DESIRED SCRIPT]", and click "Run."
7. Choose the folder containing the PowerPoint files you wish to convert.
8. The script will process each file in the folder and save them as PDFs in the same directory.

## Troubleshooting

If you encounter Error 5 at the line defining the output PDF path, ensure that the file names and folder paths do not contain invalid characters or strings that VBA might misinterpret.

## Disclaimer

Always back up your original files before running scripts or batch processes. This script is provided as-is, and while I've made efforts to ensure its accuracy and safety, I cannot be held responsible for any data loss or issues arising from its use.
