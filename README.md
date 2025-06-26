
**EXPERIMENTAL/WIP- this is not finished, and may not do what you expect! I needed this for my custom image & table plugin, but it is not yet useful as resource for those that don't have the combination of plugins I am working with.

The main reasoning behind the plugin was the need to have the .docx converter process images as something other than image tags, and process docx tables in a specific way that didn't seem feasible with the existing docx conversion logic.



A plugin for Etherpad that allows customization of how DOCX, DOC, ODT, and ODF files are transformed into HTML. This plugin extends Etherpad's native document import capabilities by providing configurable HTML output transformations. In other words, it stands on top of or replaces the default .docx converter with custom rules.

@@ -13,8 +13,9 @@ While Etherpad already supports basic DOCX-to-HTML conversion, this plugin allow

- Define specific transformations for images, tables, and other elements

- Control the exact HTML structure and attributes used in the output

- Handle embedded resources like images with custom processing


- Most importantly it serves a centralized location for multiple plugin transformations, which is to my knowledge not supported natively.