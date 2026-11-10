# Word to PDF Converter and Merger

An ASP.NET Web Forms sample application demonstrating how to convert Word documents (.doc, .docx ,.rtf, .htm, .html) to PDF and merge multiple Word and PDF files into a single unified PDF document using **Syncfusion DocIO**, **DocToPDFConverter**, and **PDF** libraries.

---

## Overview

This sample provides an intuitive web interface for common document processing workflows:

1. **Word to PDF Conversion**:
   - Convert single or multiple Microsoft  documents (.doc, .docx ,.rtf, .htm, .html) to PDF format.
   - Preserves text formatting, images, tables, shapes, and layout elements.
   - Single file conversion downloads directly as a `.pdf` file.
   - Batch conversion packages all converted PDFs into a single `.zip` archive.

2. **Merge to PDF**:
   - Combine multiple documents (.doc, .docx ,.rtf, .htm, .html) and PDF (`.pdf`) documents into a single consolidated PDF.
   - Automatically converts uploaded Word files to PDF in memory before appending them.
   - Supports reordering files using drag-and-drop or order controls prior to merging.

---

## Purpose

The primary objective of this sample is to showcase how developers can integrate Syncfusion's file format libraries into an ASP.NET Web Forms application to handle document conversion and merging without depending on Microsoft Office or Adobe Acrobat installations on the server.

Key capabilities demonstrated:
- Converting Word documents to high-fidelity PDF documents using `Syncfusion.DocToPDFConverter`.
- Merging and assembling pages from multiple PDF streams using `Syncfusion.Pdf`.
- Handling in-memory document processing and streaming generated files directly to the client response.
- Managing multi-file drag-and-drop uploads, validation, and reordering in a web interface.

---

## Project Structure

```text
Word-To-PDF-Converter-and-Merger/
│
├── README.md                                    # Project documentation
│
└── .NET FRAMEWORK/
    ├── Word-To-PDF-Converter-and-Merger.slnx    # Visual Studio Solution file
    │
    └── Word-To-PDF-Converter-and-Merger/        # ASP.NET Web Forms project
        ├── Default.aspx                         # Web UI: Tabbed interface for Convert and Merge
        ├── Default.aspx.cs                      # Code-behind: Conversion, merging, and file streaming logic
        ├── Default.aspx.designer.cs             # Designer-generated controls mapping
        │
        ├── Site.Master                          # Master page layout and common navigation
        ├── Site.Master.cs                       # Master page code-behind
        ├── Site.Master.designer.cs              # Master page designer file
        │
        ├── Global.asax                          # Application lifecycle event handlers
        ├── Global.asax.cs                       # Application initialization & routing setup
        ├── Web.config                           # Application configuration and assembly bindings
        ├── packages.config                      # NuGet package dependencies (Syncfusion, Bootstrap, etc.)
        │
        ├── App_Start/
        │   ├── BundleConfig.cs                  # Script and style bundling configuration
        │   └── RouteConfig.cs                   # Friendly URLs and routing registration
        │
        ├── Content/                             # Stylesheets (Site.css, Bootstrap CSS)
        └── Scripts/                             # Client scripts (jQuery, Bootstrap, WebForms scripts)
```

### Key Components

- **`Default.aspx`**: Contains the client-side presentation layer featuring drag-and-drop dropzones, file list previews, tab navigation, and client-side ordering for merge operations.
- **`Default.aspx.cs`**: Implements server-side processing:
  - `ConvertWordToPdfBytes()`: Converts a Word document stream into PDF bytes using `WordDocument` and `DocToPDFConverter`.
  - `btnConvertWordToPdf_Click()`: Manages single-file PDF download or multi-file ZIP generation using `ZipArchive`.
  - `btnMergeToPdf_Click()`: Processes ordered files, converts Word files to PDF in-memory, imports pages into a target `PdfDocument`, and streams the merged PDF.
- **`Web.config`**: Defines the target runtime (.NET Framework 4.8) and assembly redirects.
- **`packages.config`**: Tracks dependencies including `Syncfusion.DocIO.AspNet.Mvc5`, `Syncfusion.DocToPdfConverter.AspNet.Mvc5`, and `Syncfusion.Pdf.AspNet.Mvc5`.

---

## Prerequisites

- **Visual Studio 2022** (or Visual Studio 2019) with the **ASP.NET and web development** workload installed.
- **.NET Framework 4.8** developer pack.
- **NuGet Package Manager** (for restoring project dependencies).

---

## How to Run the Sample

1. **Clone or Download the Repository**:
   ```bash
   git clone https://github.com/your-username/Word-To-PDF-Converter-and-Merger.git
   ```

2. **Open the Solution**:
   - Navigate to the `.NET FRAMEWORK/` directory.
   - Open `Word-To-PDF-Converter-and-Merger.slnx` in Visual Studio.

3. **Restore NuGet Packages**:
   - Right-click the solution in Solution Explorer and select **Restore NuGet Packages**.

4. **Build and Run**:
   - Press **F5** or click **IIS Express** to start the application.
   - The application will launch in your default browser at `http://localhost:<port>/`.
