# Outlook PDF Converter Add-in

A Microsoft Outlook add-in that converts emails to PDF with customizable naming options and automatic categorization.

## Features

- Convert emails to PDF with local processing
- Customizable PDF naming with date, sender, and subject options
- Automatic email categorization with "PDF" category in red color
- Clean, Office UI Fabric-inspired interface
- Support for various date formats and naming separators

## Installation

1. Clone or download this repository
2. Install dependencies:
   ```bash
   npm install
   ```

3. Start the development server:
   ```bash
   npm run dev-server
   ```

4. Sideload the add-in in Outlook:
   - Open Outlook (desktop or web)
   - Go to "Get Add-ins" or "Add-ins" 
   - Choose "My add-ins" -> "Add a custom add-in" -> "Add from file"
   - Select the `manifest.xml` file from this project

## Usage

1. Open an email in Outlook
2. Click the "Convert to PDF" button in the ribbon
3. Configure your naming preferences:
   - Choose whether to include date, sender, and/or subject
   - Select date format and separator
   - Set maximum subject length
4. Preview the filename
5. Click "Convert to PDF"
6. The email will be converted to PDF and automatically categorized

## Development

### Scripts

- `npm run dev-server` - Start development server with hot reload
- `npm run build` - Build for production
- `npm run build:dev` - Build for development
- `npm run validate` - Validate the manifest file

### Project Structure

```
src/
  taskpane/
    taskpane.html    # Main UI
    taskpane.css     # Styles
    taskpane.js      # Main logic
manifest.xml         # Add-in manifest
package.json         # Dependencies and scripts
webpack.config.js    # Build configuration
```

## Technical Details

- Uses Office JavaScript API for Outlook integration
- jsPDF library for PDF generation
- Webpack for bundling and development server
- Modern JavaScript with Babel transpilation
- Responsive design for various screen sizes

## Browser Support

- Microsoft Edge
- Chrome
- Firefox  
- Safari
- Internet Explorer 11+

## License

MIT License