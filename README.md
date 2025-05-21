# ep_docx_html_customizer

A plugin for Etherpad that allows customization of how DOCX, DOC, ODT, and ODF files are transformed into HTML. This plugin extends Etherpad's native document import capabilities by providing configurable HTML output transformations. In other words, it stands on top of or replaces the default .docx converter with custom rules.

## Purpose

While Etherpad already supports basic DOCX-to-HTML conversion, this plugin allows you to:
- Customize how document elements are transformed into HTML
- Define specific transformations for images, tables, and other elements
- Control the exact HTML structure and attributes used in the output
- Handle embedded resources like images with custom processing

## Features

- Supports DOCX, DOC, ODT, and ODF file formats
- Configurable transformation rules for document elements
- Custom image processing and embedding
- Maintains compatibility with other Etherpad plugins

## Installation

## Configuration (TODO)

Add the following to your `settings.json`:

```json
{
  "ep_docx_html_customizer": {
    "transformations": [
      {
        "name": "Convert Images",
        "selector": "img",
        "action": "replaceWithSpans",
        "params": {
          "outerSpanBaseClasses": "inline-image character image-placeholder",
          "innerSpanClass": "image-inner",
          "srcClassPrefix": "image:",
          "widthClassPrefix": "image-width:",
          "heightClassPrefix": "image-height:",
          "aspectRatioClassPrefix": "imageCssAspectRatio:",
          "addResizeHandles": true,
          "addZwspPadding": true
        }
      }
      // Add more transformation rules as needed
    ]
  }
}
```

## Requirements

- LibreOffice (soffice) must be installed and configured in your Etherpad settings
- Node.js 12 or higher
- Etherpad 1.8.0 or higher

## License

Apache License 2.0 # ep_docx_html_customizer
