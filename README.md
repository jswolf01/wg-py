# Wade-Giles to Pinyin Converter

A Python tool for converting Wade-Giles romanization to Pinyin in Microsoft Word (.docx) documents.

## Features

- **Comprehensive conversion dictionary** based on standard Wade-Giles to Pinyin conversion tables
- **Handles aspirated consonants** with apostrophes (ch', k', p', t', ts')
- **Postal romanizations** support (Peking→Beijing, Nanking→Nanjing, Kwangsi→Guangxi, etc.)
- **PDF artifact handling** (converts "ii" to ü for documents converted from PDF)
- **Text box support** for documents converted from PDF (processes `<w:txbxContent>` XML elements)
- **Hyphen removal** in converted names (Tse-tung→Zedong, En-lai→Enlai)
- **Case preservation** (maintains original capitalization patterns)
- **Smart English exclusions** to avoid false positives (won't convert common English words like "to", "no", "lung")

## Installation

### Requirements

- Python 3.7 or higher
- `python-docx` library

### Setup

1. Clone or download this repository:
   ```bash
   git clone https://github.com/yourusername/wg-py.git
   cd wg-py
   ```

2. Install the required dependency:
   ```bash
   pip install python-docx
   ```

## Usage

### Basic Usage

Convert a document, automatically creating an output file with `_pinyin` suffix:
```bash
python wg_to_pinyin.py input.docx
# Creates: input_pinyin.docx
```

### Specify Output File

```bash
python wg_to_pinyin.py input.docx output.docx
```

Or using the `-o` flag:
```bash
python wg_to_pinyin.py input.docx -o output.docx
```

### Aggressive Mode

By default, common English words that happen to match Wade-Giles patterns are NOT converted to avoid false positives. Use `--aggressive` or `-a` to convert everything:

```bash
python wg_to_pinyin.py input.docx --aggressive
```

### Command-Line Help

```bash
python wg_to_pinyin.py --help
```

## Conversion Examples

### Standard Wade-Giles

| Wade-Giles | Pinyin |
|------------|--------|
| Mao Tse-tung | Mao Zedong |
| Chou En-lai | Zhou Enlai |
| Teng Hsiao-p'ing | Deng Xiaoping |
| Ch'ing Dynasty | Qing Dynasty |
| T'ang Dynasty | Tang Dynasty |
| Sung Dynasty | Song Dynasty |

### Aspirated vs. Unaspirated Consonants

| Wade-Giles | Pinyin | Notes |
|------------|--------|-------|
| ch'a | cha | Aspirated (with apostrophe) |
| cha | zha | Unaspirated (no apostrophe) |
| k'ung | kong | Aspirated |
| kung | gong | Unaspirated |
| p'ing | ping | Aspirated |
| ping | bing | Unaspirated |
| t'ai | tai | Aspirated |
| tai | dai | Unaspirated |

### Postal Romanizations

| Postal | Pinyin |
|--------|--------|
| Peking | Beijing |
| Nanking | Nanjing |
| Canton | Guangzhou |
| Kwangsi | Guangxi |
| Fukien | Fujian |
| Szechwan | Sichuan |
| Chekiang | Zhejiang |
| Tientsin | Tianjin |

### PDF Artifact Handling (ii→ü)

Documents converted from PDF often have ü rendered as "ii". The converter handles this:

| PDF Artifact | Pinyin |
|--------------|--------|
| chii | ju |
| ch'ii | qu |
| hsii | xu |
| yii | yu |
| lii | lü |

## Processing PDF-Converted Documents

If your source is a PDF, first convert it to .docx using a tool like:
- Adobe Acrobat
- Microsoft Word (Open PDF, save as .docx)
- Online converters (e.g., pdf2docx)

PDF-to-Word converters typically place text in text boxes to preserve layout. This converter handles text boxes automatically by processing the document's internal XML structure.

## Programmatic Usage

You can also use the converter as a Python module:

```python
from wg_to_pinyin import WadeGilesToPinyinConverter

# Create converter instance
converter = WadeGilesToPinyinConverter()

# Convert a document
output_path = converter.convert_docx('input.docx', 'output.docx')

# Convert with aggressive mode (converts all matches)
output_path = converter.convert_docx('input.docx', 'output.docx', aggressive=True)

# Convert text directly
text = "The Ch'ing Dynasty ruled from Peking."
converted = converter.convert_text(text)
print(converted)  # "The Qing Dynasty ruled from Beijing."
```

## Handling Edge Cases

### English Words

By default, common English words are NOT converted:
- "to" (would become "duo")
- "no" (would become "nuo")
- "lung" (would become "long")
- "tang" (would become "dang")
- "ping" (would become "bing")

However, **capitalized** versions of ambiguous words ARE converted, as they likely represent Chinese names:
- "Sung Dynasty" → "Song Dynasty"
- "Tang Dynasty" → "Dang Dynasty" (use T'ang for Tang)
- "Lung Men caves" → "Long Men caves"

### Hyphenated Names

Wade-Giles often hyphenates given names. The converter removes hyphens when converting:
- "Tse-tung" → "Zedong" (not "Ze-dong")
- "En-lai" → "Enlai"
- "Hsiao-p'ing" → "Xiaoping"

### Preserving Non-Chinese Hyphens

English hyphenated words are preserved:
- "well-known" → "well-known" (unchanged)
- "self-aware" → "self-aware" (unchanged)

## File Structure

```
wg-py/
├── wg_to_pinyin.py          # Main converter script
├── w-g to py conversion.docx # Reference conversion table
└── README.md                 # This file
```

## Limitations

- **PDF files are not directly supported.** Convert PDFs to .docx first.
- Some rare or non-standard romanizations may not be included in the conversion dictionary.
- The converter relies on pattern matching; unusual formatting or character encoding may affect results.

## Contributing

To add new conversions or postal romanizations, edit the dictionaries in `wg_to_pinyin.py`:
- `WG_TO_PINYIN` - Main Wade-Giles to Pinyin mappings
- `POSTAL_ROMANIZATIONS` - Postal/geographic romanizations
- `II_TO_UMLAUT_MAPPINGS` - PDF artifact corrections

## License

This project is provided as-is for academic and research purposes.
