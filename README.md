# PDFDocTranslator
Translate the PDF document into Japanese. The output should support either Excel or ASCII DOC format.

## How to use
Please modify the following items in config_sample.json to match your environment, and then rename the file to config.json.
- "google_api_key": "YOUR_GOOGLE_API_KEY"
- "model_name": "gemini-2.5-flash-preview-04-17"

If you are using Google Drive, please also modify the following items:
- "use_google_drive": false
- "credentials_file": "credentials.json" [^1]
[^1]: Note: Please obtain the necessary authentication information from Google and enter the filename here.
