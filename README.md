# DDMS A/R Report Converter

A simple conversion utility to create an Excel spreadsheet from a .txt DDMS A/R Report.

## Table of Contents

- [DDMS A/R Report Converter](#ddms-ar-report-converter)
  - [Table of Contents](#table-of-contents)
  - [Installation](#installation)
  - [Build](#build)
    - [Windows](#windows)
    - [Mac and Linux](#mac-and-linux)
  - [Usage](#usage)
  - [Support](#support)
  - [Contributing](#contributing)

## Installation

No installation necessary. Download the release binary and run directly.

## Build

### Windows

```
go build -ldflags -H=windowsgui

# If you have code signing set-up, run this to sign the resultant executable
signtool.exe sign /fd SHA256 ddms-ar-converter.exe
```
### Mac and Linux
Currently not supported.

## Usage

1. Run the executable
2. When prompted, select the .txt file that DDMS outputs for the Accounts Receivable Report
3. When prompted, select the location you would like to save the resultant .xlsx file

## Support

Please [open an issue](https://github.com/PercyJax/ddms-ar-converter/issues/new) for support.

## Contributing

Please contribute using [Github Flow](https://guides.github.com/introduction/flow/). Create a branch, add commits, and [open a pull request](https://github.com/PercyJax/ddms-ar-converter/compare/).
