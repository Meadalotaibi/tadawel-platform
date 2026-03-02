/**
 * PDF Export Utility
 * Uses html2canvas and jspdf to export HTML elements as styled multi-page PDFs
 */

import html2canvas from 'html2canvas';
import { jsPDF } from 'jspdf';

/**
 * Options for PDF export
 */
export interface PdfExportOptions {
  /** The HTML element to capture and export */
  element: HTMLElement;
  /** The filename for the downloaded PDF (without extension) */
  fileName: string;
  /** Optional: PDF orientation (default: 'portrait') */
  orientation?: 'portrait' | 'landscape';
  /** Optional: Scale factor for higher quality (default: 2) */
  scale?: number;
  /** Optional: Background color (default: '#0f172a' for dark theme) */
  backgroundColor?: string;
  /** Optional: Page format (default: 'a4') */
  format?: 'a4' | 'letter';
  /** Optional: Margin in mm (default: 10) */
  margin?: number;
}

/**
 * A4 dimensions in mm
 */
const A4_WIDTH_MM = 210;
const A4_HEIGHT_MM = 297;

/**
 * Exports an HTML element to a multi-page A4 PDF file
 * Captures the element using html2canvas and splits into multiple pages if needed
 * 
 * @param options - Export configuration options
 * @returns Promise that resolves when the PDF is downloaded
 * 
 * @example
 * ```typescript
 * const dashboardEl = document.getElementById('report-export-root');
 * await exportElementToPdf({
 *   element: dashboardEl,
 *   fileName: 'Saher_Weekly_Report_2024-01-01_2024-12-31'
 * });
 * ```
 */
export async function exportElementToPdf(options: PdfExportOptions): Promise<void> {
  const {
    element,
    fileName,
    orientation = 'portrait',
    scale = 2,
    backgroundColor = '#0f172a',
    margin = 10
  } = options;

  if (!element) {
    throw new Error('PDF Export: No element provided');
  }

  // Wait for charts and animations to fully render
  await new Promise(resolve => setTimeout(resolve, 300));
  
  // Wait for next animation frame to ensure all renders are complete
  await new Promise(resolve => requestAnimationFrame(resolve));

  // Capture the element as a high-resolution canvas
  const canvas = await html2canvas(element, {
    scale,
    backgroundColor,
    useCORS: true,
    allowTaint: true,
    logging: false,
    windowWidth: element.scrollWidth,
    windowHeight: element.scrollHeight
  });

  // Get canvas dimensions
  const canvasWidth = canvas.width;
  const canvasHeight = canvas.height;

  // Calculate PDF dimensions based on orientation
  const pdfWidth = orientation === 'portrait' ? A4_WIDTH_MM : A4_HEIGHT_MM;
  const pdfHeight = orientation === 'portrait' ? A4_HEIGHT_MM : A4_WIDTH_MM;
  
  // Usable area (accounting for margins)
  const usableWidth = pdfWidth - (margin * 2);
  const usableHeight = pdfHeight - (margin * 2);

  // Calculate the scale to fit canvas width to PDF usable width
  const imgScaleRatio = usableWidth / canvasWidth;
  const scaledImgWidth = usableWidth;
  const scaledImgHeight = canvasHeight * imgScaleRatio;

  // Calculate how many pages are needed
  const totalPages = Math.ceil(scaledImgHeight / usableHeight);

  // Create PDF
  const pdf = new jsPDF({
    orientation,
    unit: 'mm',
    format: 'a4'
  });

  // Height of content per page in canvas pixels
  const pageHeightInCanvasPixels = usableHeight / imgScaleRatio;

  // Add each page
  for (let page = 0; page < totalPages; page++) {
    if (page > 0) {
      pdf.addPage();
    }

    // Calculate the portion of the canvas to capture for this page
    const sourceY = page * pageHeightInCanvasPixels;
    const sourceHeight = Math.min(pageHeightInCanvasPixels, canvasHeight - sourceY);

    // Create a temporary canvas for this page's content
    const pageCanvas = document.createElement('canvas');
    pageCanvas.width = canvasWidth;
    pageCanvas.height = sourceHeight;

    const ctx = pageCanvas.getContext('2d');
    if (ctx) {
      // Draw the portion of the original canvas onto the page canvas
      ctx.drawImage(
        canvas,
        0, sourceY,                    // Source x, y
        canvasWidth, sourceHeight,     // Source width, height
        0, 0,                          // Destination x, y
        canvasWidth, sourceHeight      // Destination width, height
      );
    }

    // Convert page canvas to image data
    const pageImgData = pageCanvas.toDataURL('image/png', 1.0);

    // Calculate the height for this page's image in PDF units
    const pageImgHeight = sourceHeight * imgScaleRatio;

    // Add the image to the PDF page
    pdf.addImage(
      pageImgData,
      'PNG',
      margin,
      margin,
      scaledImgWidth,
      pageImgHeight,
      undefined,
      'FAST'
    );
  }

  // Download the PDF
  pdf.save(`${fileName}.pdf`);
}

/**
 * Exports an HTML element to a single-page PDF (fits content to one page)
 * Useful for smaller content that doesn't need pagination
 * 
 * @param options - Export configuration options
 * @returns Promise that resolves when the PDF is downloaded
 */
export async function exportElementToSinglePagePdf(options: PdfExportOptions): Promise<void> {
  const {
    element,
    fileName,
    orientation = 'landscape',
    scale = 2,
    backgroundColor = '#0f172a'
  } = options;

  if (!element) {
    throw new Error('PDF Export: No element provided');
  }

  // Wait for charts and animations to fully render
  await new Promise(resolve => setTimeout(resolve, 300));
  await new Promise(resolve => requestAnimationFrame(resolve));

  // Capture the element as a canvas
  const canvas = await html2canvas(element, {
    scale,
    backgroundColor,
    useCORS: true,
    allowTaint: true,
    logging: false
  });

  // Get canvas dimensions
  const imgWidth = canvas.width;
  const imgHeight = canvas.height;

  // Create PDF with dimensions matching the canvas (scaled appropriately)
  const pdf = new jsPDF({
    orientation,
    unit: 'px',
    format: [imgWidth, imgHeight]
  });

  // Add the canvas image to the PDF
  const imgData = canvas.toDataURL('image/png');
  pdf.addImage(imgData, 'PNG', 0, 0, imgWidth, imgHeight);

  // Download the PDF
  pdf.save(`${fileName}.pdf`);
}
