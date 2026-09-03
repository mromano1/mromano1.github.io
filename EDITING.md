# Editing the Project Management Lessons Site

This version is designed so you can edit it directly in GitHub without rebuilding the PowerPoint conversion.

## Edit text on a page
1. Open your GitHub repository.
2. Click the HTML file you want to change, such as `index.html` or `lesson-3.html`.
3. Click the pencil icon (**Edit this file**).
4. Make your change.
5. Click **Commit changes**. GitHub Pages will republish the site automatically.

## Add or replace a downloadable resource
The downloadable files are in the `resources` folder. You can upload a new Word, Excel, PDF, or other file there.

To add a link, copy an existing link in `index.html` and change the filename and visible label. A typical link looks like:

```html
<a class="resource-link" href="resources/Your-File.docx" target="_blank" rel="noopener noreferrer">Your Link Name ↗</a>
```

The `target="_blank"` setting opens the link in a new browser tab/window. Office files may download instead of displaying, depending on the visitor's browser.

## Edit lesson content
Each lesson is a separate HTML file (`lesson-1.html` through `lesson-8.html`). The visible slide text and speaker notes are stored near the bottom of each lesson file in `window.SLIDES=[...]`. You can edit text there using GitHub's editor.

For simple edits, avoid changing punctuation that is part of the JavaScript structure (quotation marks, commas, brackets). If you want major lesson revisions, it is safer to update the PowerPoint and regenerate that lesson page.

## Change appearance
Colors, spacing, fonts, buttons, and resource-link styling are in `assets/style.css`.
