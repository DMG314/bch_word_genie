# Daniel was here
#


# Word Title Page Add-in

A Microsoft Word add-in that creates formatted title pages with user-provided information.

## Features

- Form interface with fields for Title, Lab Name, Professor Name, and Date
- Inserts a professionally formatted title page at the beginning of the document
- Centers content horizontally and applies heading styles
- Clean, responsive user interface

## Setup Instructions

1. Install dependencies:
   ```bash
   npm install
   ```

2. Start the development server:
   ```bash
   npm start
   ```

3. The add-in will be served at `https://localhost:3000`

## Installing the Add-in

1. Open Microsoft Word
2. Go to Insert > Add-ins > My Add-ins
3. Click "Upload My Add-in"
4. Select the `manifest.xml` file from this project
5. The add-in will appear in the Home tab as "Create Title Page"

## Usage

1. Click the "Create Title Page" button in the Word ribbon
2. Fill out the form with:
   - Title: The main title of your document
   - Lab Name: Name of the laboratory or course
   - Professor Name: Name of the instructor
   - Date: Date of the document
3. Click "Insert" to add the formatted title page

The title page will be inserted at the beginning of your document with:
- Title formatted as Heading 1
- Lab Name formatted as Heading 2
- Professor Name formatted as Heading 3
- Date formatted as Heading 4
- All content centered horizontally

## Files Structure

- `manifest.xml` - Add-in manifest file
- `taskpane.html` - Main HTML interface
- `taskpane.js` - JavaScript functionality
- `taskpane.css` - Styling for the interface
- `commands.html` - Commands page (required by Office)
- `package.json` - Node.js dependencies and scripts

## Development

To make changes to the add-in:

1. Edit the HTML, CSS, or JavaScript files
2. Refresh the task pane in Word to see changes
3. Use `npm run validate` to check the manifest file

## Troubleshooting

- Ensure you're using HTTPS (required for Office Add-ins)
- Check that the manifest.xml file is valid
- Verify that the development server is running on port 3000
- Make sure Word is connected to the internet for add-in functionality
