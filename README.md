# ep_docx_html_customizer


**EXPERIMENTAL - this is not finished and I'm not sure if it works! I seemed to need this for my custom image & table plugin!
The main reasoning behind the plugin was the need to have the .docx converter process images as something other than image tags, and process docx tables in a specific way that didn't seem feasible with the existing docx conversion logic.

In other words, it is not customizable yet. Right now I just have the basics for transforming two specific HTML elements. Later, I will try to make it support arbitrary definitions through the Etherpad settings.json

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
