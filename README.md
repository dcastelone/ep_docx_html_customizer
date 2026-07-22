# ep_docx_html_customizer

Normalizes office-document imports and rich clipboard content for Etherpad installations that use structured tables, hyperlinks, inline images, colors, and related formatting plugins.

## Supported imports

- Microsoft Word (`.docx`, `.doc`)
- OpenDocument Text (`.odt`, `.odf`)

The plugin converts supported documents through LibreOffice, transforms the resulting HTML into Etherpad-compatible structures, and lets unsupported file types continue through Etherpad's normal import pipeline.

## Requirements

- Etherpad 3.3.2 or a later 3.x release
- LibreOffice installed on the Etherpad host or image
- Etherpad's `soffice` setting configured with the LibreOffice executable path
- `ep_images_extended` configured for `s3_presigned` storage when imported or pasted images must be retained

Example Etherpad settings:

```json
{
  "soffice": "/usr/bin/soffice",
  "ep_images_extended": {
    "storage": {
      "type": "s3_presigned",
      "region": "us-east-1",
      "bucket": "example-images",
      "publicURL": "https://cdn.example.org/images/",
      "expires": 900
    },
    "fileTypes": ["png", "jpg", "jpeg", "webp", "gif"],
    "maxFileSize": 10485760
  }
}
```

The AWS SDK uses its normal credential provider chain. In AWS, prefer a task or instance role over long-lived access keys.

## Image handling

Document images are uploaded to the S3 storage configured for `ep_images_extended`. Rich clipboard content can use the authenticated same-origin image proxy before upload. The proxy accepts only HTTP(S) URLs, blocks common loopback and private-address targets, limits requests by client IP, enforces fetch time and size limits, and requires an Etherpad session or pad cookie.

If S3 image storage is unavailable, image retention fails closed instead of embedding an unexpected external or local source.

## Companion-plugin compatibility

The transformations are designed for content used with `ep_data_tables`, `ep_hyperlinked_text`, `ep_images_extended`, and compatible font-color plugins. Test the complete installed plugin set against representative source documents before production rollout.

## Installation

From the Etherpad directory:

```sh
pnpm run plugins i ep_docx_html_customizer
```

Restart Etherpad after installation.

## Development

```sh
pnpm install --frozen-lockfile
pnpm test
```

Licensed under the Apache License 2.0.
