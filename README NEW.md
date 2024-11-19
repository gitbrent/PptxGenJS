# PptxGenJS-Appinio

this is a Node.js demo for PptxGenJS (a PowerPoint presentation generation library). Here's how you can run it locally:

1. First, make sure you have Node.js installed on your system.

2. Install the required dependency:
```bash
npm install pptxgenjs
```

3. You can run the file in several ways (as mentioned in the file header):
```bash
# Run local tests with callbacks
node demo.js

# Run all pre-defined tests
node demo.js All

# Run a specific test (e.g., Text)
node demo.js Text
```

When you run the script:
- It will create a PowerPoint presentation named "PptxGenJS_Demo_Node"
- The file will be saved in your current working directory
- The script will create multiple slides with different charts (funnel charts and waterfall charts)
- You'll see console output showing the export status and any errors

The script will show you:
- The pptxgenjs version being used
- The save location (current working directory)
- Export confirmation messages
- A BASE64 preview of the generated file

If you're having trouble with the imports, make sure your project structure matches the import paths:
```
your-project/
├── dist/
│   └── pptxgen.cjs.js
├── modules/
│   └── demos.mjs
└── demo.js
```