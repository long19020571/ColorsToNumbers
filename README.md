All Rights Reserved © 2025 Nguyen Viet Long

This project and all associated files are the intellectual property of Nguyen Viet Long.

You are not allowed to use, copy, modify, distribute, or exploit this project in any form without explicit written permission from the author.

For inquiries about usage, licensing, or collaboration, please contact:
nguyenvietlong1201@gmail.com

# ColorsToNumbers

## Overview

**ColorsToNumbers** is an automation tool written in C# for analyzing, indexing, and reporting colored polygons in Adobe Illustrator files. The application scans all filled polygons, assigns a unique index to each color, labels each region, calculates area statistics, and exports high-quality PNG images and comprehensive PDF reports.

---

## Features

- **Automatic Color Indexing:** Scans all filled polygons and assigns a unique index to each distinct color.
- **Polygon Labeling:** Places index labels at the centroid of each polygon for clear identification.
- **Legend & Area Calculation:** Generates a color legend with area statistics for each color.
- **Batch Export:** Exports labeled artwork as high-resolution PNG images and compiles them into organized PDF reports.
- **Progress Feedback:** Provides real-time progress updates in the console for each processing stage.
- **Parallel Processing:** Utilizes parallelism for efficient handling of large and complex documents.

---

## Project Structure

- **ColorIndex:** Manages color-to-index mapping and color properties.
- **PolygonInfo:** Stores polygon geometry and its assigned color index.
- **ColorsToNumber2:** Main class for document processing, labeling, exporting, and reporting.

---

## Technologies & Innovations

- **.NET 8 & C# 12:** Utilizes the latest language features and performance improvements.
- **Adobe Illustrator COM Interop:** Automates Illustrator via its COM Interop.
- **NetTopologySuite:** Advanced geometry processing for accurate polygon analysis and manipulation.
- **PdfSharp:** High-quality PDF generation, including custom graphics and text rendering.
- **Parallel Programming:** Uses `Parallel.ForEach` and asynchronous tasks for fast, scalable processing.
- **Modern CLI Feedback:** Real-time console progress bars and status messages.

---

## Application Scenarios

- **Design Documentation:** Automatically generate labeled diagrams and color legends for design handoff or archiving.
- **Production Preparation:** Create clear, indexed artwork for manufacturing, printing, or embroidery.
- **Quality Control:** Analyze color usage and area distribution in complex vector graphics.

---

## Requirements

- **.NET 8 SDK**
- **Adobe Illustrator 2025**
- **NuGet Packages:**
  - NetTopologySuite
  - PdfSharp
- **Font:** Arial (for PDF output)

---

## Usage

1. Open your Illustrator document.
2. Run the application.
3. Check the output in the `/MyFiles` directory (images, processed AI files) and the generated PDF reports.

---

## License

MIT License (or specify your license here)

---

**Contact:**  
Please open an issue or pull request if you have questions or want to contribute.
