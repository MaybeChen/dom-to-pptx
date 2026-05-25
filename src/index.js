// src/index.js
import * as PptxGenJSImport from 'pptxgenjs';
import html2canvas from 'html2canvas';
import { PPTXEmbedFonts } from './font-embedder.js';
import { normalizePptxZip } from './pptx-normalizer.js';
import JSZip from 'jszip';

// Normalize import
const PptxGenJS = PptxGenJSImport?.default ?? PptxGenJSImport;

import {
  parseColor,
  getTextStyle,
  isTextContainer,
  getVisibleShadow,
  generateGradientSVG,
  getRotation,
  getWritingModeVert,
  svgToPng,
  svgToSvg,
  getPadding,
  getSoftEdges,
  generateBlurredSVG,
  getBorderInfo,
  generateCompositeBorderSVG,
  isClippedByParent,
  generateCustomShapeSVG,
  getUsedFontFamilies,
  getAutoDetectedFonts,
  extractTableData,
  collectTextParts,
} from './utils.js';
import { getProcessedImage } from './image-processor.js';

const PPI = 96;
const PX_TO_INCH = 1 / PPI;

/**
 * Main export function.
 * @param {HTMLElement | string | Array<HTMLElement | string>} target
 * @param {Object} options
 * @param {string} [options.fileName]
 * @param {boolean} [options.skipDownload=false] - If true, prevents automatic download
 * @param {Object} [options.listConfig] - Config for bullets
 * @param {boolean} [options.svgAsVector=false] - If true, keeps SVG as vector (for Convert to Shape in PowerPoint)
 * @param {boolean} [options.skipNormalize=false] - If true, skips re-zipping with DEFLATE
 *   and stripping dangling [Content_Types].xml Overrides. Leave it false unless you are
 *   debugging the raw PptxGenJS output, otherwise Microsoft PowerPoint may reject the file.
 * @returns {Promise<Blob>} - Returns the generated PPTX Blob
 */
export async function exportToPptx(target, options = {}) {
  const resolvePptxConstructor = (pkg) => {
    if (!pkg) return null;
    if (typeof pkg === 'function') return pkg;
    if (pkg && typeof pkg.default === 'function') return pkg.default;
    if (pkg && typeof pkg.PptxGenJS === 'function') return pkg.PptxGenJS;
    if (pkg && pkg.PptxGenJS && typeof pkg.PptxGenJS.default === 'function')
      return pkg.PptxGenJS.default;
    return null;
  };

  const PptxConstructor = resolvePptxConstructor(PptxGenJS);
  if (!PptxConstructor) throw new Error('PptxGenJS constructor not found.');
  const pptx = new PptxConstructor();

  // 1. Layout Handling
  let finalWidth = 10; // default 16:9
  let finalHeight = 5.625;

  if (options.width && options.height) {
    pptx.defineLayout({ name: 'CUSTOM', width: options.width, height: options.height });
    pptx.layout = 'CUSTOM';
    finalWidth = options.width;
    finalHeight = options.height;
  } else if (options.layout) {
    pptx.layout = options.layout;
    // Map standard layouts for internal scale calculation if possible,
    // though PptxGenJS defaults to 16:9 if unknown.
    if (options.layout === 'LAYOUT_4x3') {
      finalWidth = 10;
      finalHeight = 7.5;
    } else if (options.layout === 'LAYOUT_16x10') {
      finalWidth = 10;
      finalHeight = 6.25;
    } else if (options.layout === 'LAYOUT_WIDE') {
      finalWidth = 13.3;
      finalHeight = 7.5;
    }
  } else {
    pptx.layout = 'LAYOUT_16x9';
  }

  // Pass these dimensions to options so processSlide can use them
  const extendedOptions = {
    ...options,
    _slideWidth: finalWidth,
    _slideHeight: finalHeight,
  };

  const elements = Array.isArray(target) ? target : [target];

  for (const el of elements) {
    const root = typeof el === 'string' ? document.querySelector(el) : el;
    if (!root) {
      console.warn('Element not found, skipping slide:', el);
      continue;
    }
    const slide = pptx.addSlide();
    await processSlide(root, slide, pptx, extendedOptions);
  }

  // 3. Font Embedding Logic
  let finalBlob;
  let fontsToEmbed = options.fonts || [];

  if (options.autoEmbedFonts) {
    // A. Scan DOM for used font families
    const usedFamilies = getUsedFontFamilies(elements);

    // B. Scan CSS for URLs matches
    const detectedFonts = await getAutoDetectedFonts(usedFamilies);

    // C. Merge (Avoid duplicates)
    const explicitNames = new Set(fontsToEmbed.map((f) => f.name));
    for (const autoFont of detectedFonts) {
      if (!explicitNames.has(autoFont.name)) {
        fontsToEmbed.push(autoFont);
      }
    }

    if (detectedFonts.length > 0) {
      console.log(
        'Auto-detected fonts:',
        detectedFonts.map((f) => f.name)
      );
    }
  }

  if (fontsToEmbed.length > 0) {
    // Generate initial PPTX
    const initialBlob = await pptx.write({ outputType: 'blob' });

    // Load into Embedder
    const zip = await JSZip.loadAsync(initialBlob);
    const embedder = new PPTXEmbedFonts();
    await embedder.loadZip(zip);

    // Fetch and Embed
    for (const fontCfg of fontsToEmbed) {
      try {
        const response = await fetch(fontCfg.url);
        if (!response.ok) throw new Error(`Failed to fetch ${fontCfg.url}`);
        const buffer = await response.arrayBuffer();

        // Infer type
        const ext = fontCfg.url.split('.').pop().split(/[?#]/)[0].toLowerCase();
        let type = 'ttf';
        if (['woff', 'otf'].includes(ext)) type = ext;

        await embedder.addFont(fontCfg.name, buffer, type);
      } catch (e) {
        console.warn(`Failed to embed font: ${fontCfg.name} (${fontCfg.url})`, e);
      }
    }

    await embedder.updateFiles();
    if (options.skipNormalize !== true) {
      await normalizePptxZip(zip);
    }
    finalBlob = await embedder.generateBlob();
  } else {
    // No fonts to embed — still re-zip with DEFLATE and strip dangling Overrides
    // so Microsoft PowerPoint accepts the file (PptxGenJS leaves both issues
    // unresolved on its own; see 错误诊断.md).
    const initialBlob = await pptx.write({ outputType: 'blob' });
    if (options.skipNormalize === true) {
      finalBlob = initialBlob;
    } else {
      const zip = await JSZip.loadAsync(initialBlob);
      await normalizePptxZip(zip);
      finalBlob = await zip.generateAsync({
        type: 'blob',
        compression: 'DEFLATE',
        compressionOptions: { level: 6 },
      });
    }
  }

  // 4. Output Handling
  // If skipDownload is NOT true, proceed with browser download
  if (!options.skipDownload) {
    const fileName = options.fileName || 'export.pptx';
    const url = URL.createObjectURL(finalBlob);
    const a = document.createElement('a');
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  // Always return the blob so the caller can use it (e.g. upload to server)
  return finalBlob;
}

/**
 * Worker function to process a single DOM element into a single PPTX slide.
 * @param {HTMLElement} root - The root element for this slide.
 * @param {PptxGenJS.Slide} slide - The PPTX slide object to add content to.
 * @param {PptxGenJS} pptx - The main PPTX instance.
 */
async function processSlide(root, slide, pptx, globalOptions = {}) {
  const rootRect = root.getBoundingClientRect();
  const PPTX_WIDTH_IN = globalOptions._slideWidth || 10;
  const PPTX_HEIGHT_IN = globalOptions._slideHeight || 5.625;

  const contentWidthIn = rootRect.width * PX_TO_INCH;
  const contentHeightIn = rootRect.height * PX_TO_INCH;
  const scale = Math.min(PPTX_WIDTH_IN / contentWidthIn, PPTX_HEIGHT_IN / contentHeightIn);

  const layoutConfig = {
    rootX: rootRect.x,
    rootY: rootRect.y,
    scale: scale,
    offX: (PPTX_WIDTH_IN - contentWidthIn * scale) / 2,
    offY: (PPTX_HEIGHT_IN - contentHeightIn * scale) / 2,
  };

  const renderQueue = [];
  const asyncTasks = []; // Queue for heavy operations (Images, Canvas)
  let domOrderCounter = 0;

  // Sync Traversal Function
  function collect(node, parentZIndex, parentOpacity = 1) {
    const order = domOrderCounter++;

    let currentZ = parentZIndex;
    let currentOpacity = parentOpacity;
    let nodeStyle = null;
    const nodeType = node.nodeType;

    if (nodeType === 1) {
      nodeStyle = window.getComputedStyle(node);
      const elOpacity = parseFloat(nodeStyle.opacity);
      if (!isNaN(elOpacity)) {
        currentOpacity *= elOpacity;
      }

      // Optimization: Skip completely hidden elements immediately
      if (
        nodeStyle.display === 'none' ||
        nodeStyle.visibility === 'hidden' ||
        currentOpacity === 0
      ) {
        return;
      }
      if (nodeStyle.zIndex !== 'auto') {
        currentZ = parseInt(nodeStyle.zIndex);
      }
    }

    // Prepare the item. If it needs async work, it returns a 'job'
    const result = prepareRenderItem(
      node,
      { ...layoutConfig, root },
      order,
      pptx,
      currentZ,
      nodeStyle,
      { ...globalOptions, _inheritedOpacity: parentOpacity }
    );

    if (result) {
      if (result.items) {
        // Push items immediately to queue (data might be missing but filled later)
        renderQueue.push(...result.items);
      }
      if (result.job) {
        // Push the promise-returning function to the task list
        asyncTasks.push(result.job);
      }
      if (result.stopRecursion) return;
    }

    // Recurse children synchronously
    const childNodes = node.childNodes;
    for (let i = 0; i < childNodes.length; i++) {
      collect(childNodes[i], currentZ, currentOpacity);
    }
  }

  // 1. Traverse and build the structure (Fast)
  collect(root, 0);

  // 2. Execute all heavy tasks in parallel (Fast)
  if (asyncTasks.length > 0) {
    await Promise.all(asyncTasks.map((task) => task()));
  }

  // 3. Cleanup and Sort
  // Remove items that failed to generate data (marked with skip)
  const finalQueue = renderQueue.filter(
    (item) => !item.skip && (item.type !== 'image' || item.options.data)
  );

  finalQueue.sort((a, b) => {
    if (a.zIndex !== b.zIndex) return a.zIndex - b.zIndex;
    return a.domOrder - b.domOrder;
  });

  // 4. Add to Slide
  for (const item of finalQueue) {
    if (item.type === 'shape') slide.addShape(item.shapeType, item.options);
    if (item.type === 'image') slide.addImage(item.options);
    if (item.type === 'text') slide.addText(item.textParts, item.options);
    if (item.type === 'table') {
      slide.addTable(item.tableData.rows, {
        x: item.options.x,
        y: item.options.y,
        w: item.options.w,
        colW: item.tableData.colWidths, // Essential for correct layout
        autoPage: false,
        // Remove default table styles so our extracted CSS applies cleanly
        border: { type: 'none' },
        fill: { color: 'FFFFFF', transparency: 100 },
      });
    }
  }
}

/**
 * Optimized html2canvas wrapper
 * Includes fix for cropped icons by adjusting styles in the cloned document.
 */
async function elementToCanvasImage(node, widthPx, heightPx) {
  return new Promise((resolve) => {
    // 1. Assign a temp ID to locate the node inside the cloned document
    const originalId = node.id;
    const tempId = 'pptx-capture-' + Math.random().toString(36).substr(2, 9);
    node.id = tempId;

    const width = Math.max(Math.ceil(widthPx), 1);
    const height = Math.max(Math.ceil(heightPx), 1);
    const style = window.getComputedStyle(node);

    // Add padding to the clone to capture spilling content (like extensive font glyphs)
    const padding = 10;

    html2canvas(node, {
      backgroundColor: null,
      logging: false,
      scale: 3, // Higher scale for sharper icons
      useCORS: true, // critical for external fonts/images
      width: width + padding * 2, // Capture a larger area
      height: height + padding * 2,
      x: -padding, // Offset capture to include the padding
      y: -padding,
      onclone: (clonedDoc) => {
        const clonedNode = clonedDoc.getElementById(tempId);
        if (clonedNode) {
          // --- FIX: CLIP & FONT ISSUES ---
          // Apply styles DIRECTLY to elements to ensure html2canvas picks them up
          // This avoids issues where <style> tags in onclone are ignored or delayed

          // 1. Force FontAwesome Family on Icons
          const icons = clonedNode.querySelectorAll('.fa, .fas, .far, .fab');
          icons.forEach((icon) => {
            icon.style.setProperty('font-family', 'FontAwesome', 'important');
          });

          // 2. Fix Image Display
          const images = clonedNode.querySelectorAll('img');
          images.forEach((img) => {
            img.style.setProperty('display', 'inline-block', 'important');
          });

          // 3. Force overflow visible on the container so glyphs bleeding out aren't cut
          clonedNode.style.overflow = 'visible';

          // 4. Adjust alignment for Icons to prevent baseline clipping
          // (Applies to <i>, <span>, or standard icon classes)
          const tag = clonedNode.tagName;
          if (tag === 'I' || tag === 'SPAN' || clonedNode.className.includes('fa-')) {
            // Flex center helps align the glyph exactly in the middle of the box
            // preventing top/bottom cropping due to line-height mismatches.
            clonedNode.style.display = 'inline-flex';
            clonedNode.style.justifyContent = 'center';
            clonedNode.style.alignItems = 'center';
            clonedNode.style.setProperty('font-family', 'FontAwesome', 'important'); // Ensure root icon gets it too

            // Remove margins that might offset the capture
            clonedNode.style.margin = '0';

            // Ensure the font fits
            clonedNode.style.lineHeight = '1';
            clonedNode.style.verticalAlign = 'middle';
          }
        }
      },
    })
      .then((canvas) => {
        // Restore the original ID
        if (originalId) node.id = originalId;
        else node.removeAttribute('id');

        const destCanvas = document.createElement('canvas');
        destCanvas.width = width;
        destCanvas.height = height;
        const ctx = destCanvas.getContext('2d');

        // Draw captured canvas (which is padded) back to the original size
        // We need to draw the CENTER of the source canvas to the destination
        // The source canvas is (width + 2*padding) * scale
        // We want to draw the crop starting at padding*scale
        const scale = 3;
        const sX = padding * scale;
        const sY = padding * scale;
        const sW = width * scale;
        const sH = height * scale;

        ctx.drawImage(canvas, sX, sY, sW, sH, 0, 0, width, height);

        // --- Border Radius Clipping (Existing Logic) ---
        let tl = parseFloat(style.borderTopLeftRadius) || 0;
        let tr = parseFloat(style.borderTopRightRadius) || 0;
        let br = parseFloat(style.borderBottomRightRadius) || 0;
        let bl = parseFloat(style.borderBottomLeftRadius) || 0;

        const f = Math.min(
          width / (tl + tr) || Infinity,
          height / (tr + br) || Infinity,
          width / (br + bl) || Infinity,
          height / (bl + tl) || Infinity
        );

        if (f < 1) {
          tl *= f;
          tr *= f;
          br *= f;
          bl *= f;
        }

        if (tl + tr + br + bl > 0) {
          ctx.globalCompositeOperation = 'destination-in';
          ctx.beginPath();
          ctx.moveTo(tl, 0);
          ctx.lineTo(width - tr, 0);
          ctx.arcTo(width, 0, width, tr, tr);
          ctx.lineTo(width, height - br);
          ctx.arcTo(width, height, width - br, height, br);
          ctx.lineTo(bl, height);
          ctx.arcTo(0, height, 0, height - bl, bl);
          ctx.lineTo(0, tl);
          ctx.arcTo(0, 0, tl, 0, tl);
          ctx.closePath();
          ctx.fill();
        }

        resolve(destCanvas.toDataURL('image/png'));
      })
      .catch((e) => {
        if (originalId) node.id = originalId;
        else node.removeAttribute('id');
        console.warn('Canvas capture failed for node', node, e);
        resolve(null);
      });
  });
}

/**
 * Helper to identify elements that should be rendered as icons (Images).
 * Detects Custom Elements AND generic tags (<i>, <span>) with icon classes/pseudo-elements.
 */
function isIconElement(node) {
  // 1. Custom Elements (hyphenated tags) or Explicit Library Tags
  const tag = node.tagName.toUpperCase();
  if (
    tag.includes('-') ||
    [
      'MATERIAL-ICON',
      'ICONIFY-ICON',
      'REMIX-ICON',
      'ION-ICON',
      'EVA-ICON',
      'BOX-ICON',
      'FA-ICON',
    ].includes(tag)
  ) {
    return true;
  }

  // 2. Class-based Icons (FontAwesome, Bootstrap, Material symbols) on <i> or <span>
  if (tag === 'I' || tag === 'SPAN') {
    const cls = node.getAttribute('class') || '';
    if (
      typeof cls === 'string' &&
      (cls.includes('fa-') ||
        cls.includes('fas') ||
        cls.includes('far') ||
        cls.includes('fab') ||
        cls.includes('bi-') ||
        cls.includes('material-icons') ||
        cls.includes('icon'))
    ) {
      // Double-check: Must have pseudo-element content to be a CSS icon
      const before = window.getComputedStyle(node, '::before').content;
      const after = window.getComputedStyle(node, '::after').content;
      const hasContent = (c) => c && c !== 'none' && c !== 'normal' && c !== '""';

      if (hasContent(before) || hasContent(after)) return true;
    }
  }

  return false;
}

/**
 * Replaces createRenderItem.
 * Returns { items: [], job: () => Promise, stopRecursion: boolean }
 */
function prepareRenderItem(
  node,
  config,
  domOrder,
  pptx,
  effectiveZIndex,
  computedStyle,
  globalOptions = {}
) {
  // 1. Text Node Handling
  if (node.nodeType === 3) {
    const textContent = node.nodeValue.trim();
    if (!textContent) return null;

    const parent = node.parentElement;
    if (!parent) return null;

    if (isTextContainer(parent)) return null; // Parent handles it

    const range = document.createRange();
    range.selectNode(node);
    const rect = range.getBoundingClientRect();
    range.detach();

    const style = window.getComputedStyle(parent);
    const widthPx = rect.width;
    const heightPx = rect.height;
    const unrotatedW = widthPx * PX_TO_INCH * config.scale;
    const unrotatedH = heightPx * PX_TO_INCH * config.scale;

    const x = config.offX + (rect.left - config.rootX) * PX_TO_INCH * config.scale;
    const y = config.offY + (rect.top - config.rootY) * PX_TO_INCH * config.scale;

    return {
      items: [
        {
          type: 'text',
          zIndex: effectiveZIndex,
          domOrder,
          textParts: [
            {
              text: textContent,
              options: getTextStyle(
                style,
                config.scale,
                true,
                globalOptions._inheritedOpacity || 1
              ),
            },
          ],
          // Honor CSS white-space: a `nowrap`/`pre` element must not re-wrap in
          // the exported slide (otherwise a single line measured in the browser
          // can wrap in PowerPoint/LibreOffice due to font-metric differences).
          options: {
            x,
            y,
            w: unrotatedW,
            h: unrotatedH,
            margin: 0,
            autoFit: true,
            wrap: !(style.whiteSpace === 'nowrap' || style.whiteSpace === 'pre'),
          },
        },
      ],
      stopRecursion: false,
    };
  }

  if (node.nodeType !== 1) return null;
  const style = computedStyle; // Use pre-computed style

  const rect = node.getBoundingClientRect();
  if (rect.width < 0.5 || rect.height < 0.5) return null;

  const zIndex = effectiveZIndex;
  const rotation = getRotation(style.transform);
  const writingModeVert = getWritingModeVert(style.writingMode, style.textOrientation);
  const elementOpacity = parseFloat(style.opacity);
  const localOpacity = isNaN(elementOpacity) ? 1 : elementOpacity;
  const inheritedOpacity = globalOptions._inheritedOpacity || 1;
  const safeOpacity = localOpacity * inheritedOpacity;

  // Prefer the sub-pixel rect size to avoid 1px text-wrap artifacts caused by
  // offsetWidth/offsetHeight being integer-rounded. When the element is rotated
  // we must fall back to offset* because rect.* describes the rotated bounding box.
  const widthPx = rotation === 0 ? rect.width || node.offsetWidth : node.offsetWidth || rect.width;
  const heightPx =
    rotation === 0 ? rect.height || node.offsetHeight : node.offsetHeight || rect.height;
  const unrotatedW = widthPx * PX_TO_INCH * config.scale;
  const unrotatedH = heightPx * PX_TO_INCH * config.scale;
  const centerX = rect.left + rect.width / 2;
  const centerY = rect.top + rect.height / 2;

  let x = config.offX + (centerX - config.rootX) * PX_TO_INCH * config.scale - unrotatedW / 2;
  let y = config.offY + (centerY - config.rootY) * PX_TO_INCH * config.scale - unrotatedH / 2;
  let w = unrotatedW;
  let h = unrotatedH;

  const items = [];

  if (node.tagName === 'TABLE') {
    const tableData = extractTableData(node, config.scale);
    const tableItems = [
      {
        type: 'table',
        zIndex: effectiveZIndex,
        domOrder,
        tableData: tableData,
        options: { x, y, w: unrotatedW, h: unrotatedH },
      },
    ];

    // 1. Check for Background / Shadow / Radius on the table itself
    const shadowStr = style.boxShadow;
    const hasShadow = shadowStr && shadowStr !== 'none';
    const borderRadius = parseFloat(style.borderRadius) || 0;
    const bgColor = parseColor(style.backgroundColor);
    const hasBg = bgColor.hex && bgColor.opacity > 0;

    if (hasShadow || borderRadius > 0 || hasBg) {
      const transparency = (1 - bgColor.opacity) * 100;
      const shadow = hasShadow ? getVisibleShadow(shadowStr, config.scale) : null;
      let shapeType = pptx.ShapeType.rect;
      let rectRadius = 0;

      if (borderRadius > 0) {
        shapeType = pptx.ShapeType.roundRect;
        let cappedRadiusPx = Math.min(borderRadius, Math.min(widthPx, heightPx) / 2);
        rectRadius = cappedRadiusPx * PX_TO_INCH * config.scale;
      }

      // Add a backing shape item before the table
      tableItems.unshift({
        type: 'shape',
        zIndex: effectiveZIndex,
        domOrder, // Same domOrder ensures it renders before the table (queue order)
        shapeType,
        options: {
          x,
          y,
          w: unrotatedW,
          h: unrotatedH,
          fill: hasBg ? { color: bgColor.hex, transparency } : { type: 'none' },
          shadow,
          rectRadius,
        },
      });
    }

    return {
      items: tableItems,
      stopRecursion: true,
    };
  }

  if ((node.tagName === 'UL' || node.tagName === 'OL') && !isComplexHierarchy(node)) {
    const listItems = [];
    const liChildren = Array.from(node.children).filter((c) => c.tagName === 'LI');

    liChildren.forEach((child, index) => {
      const liStyle = window.getComputedStyle(child);
      const liRect = child.getBoundingClientRect();
      const parentRect = node.getBoundingClientRect(); // node is UL/OL

      // 1. Determine Bullet Config
      let bullet = { type: 'bullet' };
      const listStyleType = liStyle.listStyleType || 'disc';

      if (node.tagName === 'OL' || listStyleType === 'decimal') {
        bullet = { type: 'number' };
      } else if (listStyleType === 'none') {
        bullet = false;
      } else {
        let code = '2022'; // disc
        if (listStyleType === 'circle') code = '25CB';
        if (listStyleType === 'square') code = '25A0';

        // --- CHANGE: Color & Size Logic (Option > ::marker > CSS color) ---
        let finalHex = '000000';
        let markerFontSize = null;

        // A. Check Global Option override
        if (globalOptions?.listConfig?.color) {
          finalHex = parseColor(globalOptions.listConfig.color).hex || '000000';
        }
        // B. Check ::marker pseudo element (supported in modern browsers)
        else {
          const markerStyle = window.getComputedStyle(child, '::marker');
          const markerColor = parseColor(markerStyle.color);
          if (markerColor.hex) {
            finalHex = markerColor.hex;
          } else {
            // C. Fallback to LI text color
            const colorObj = parseColor(liStyle.color);
            if (colorObj.hex) finalHex = colorObj.hex;
          }

          // Check ::marker font-size
          const markerFs = parseFloat(markerStyle.fontSize);
          if (!isNaN(markerFs) && markerFs > 0) {
            // Convert px->pt for PPTX
            markerFontSize = markerFs * 0.75 * config.scale;
          }
        }

        bullet = { code, color: finalHex };
        if (markerFontSize) {
          bullet.fontSize = markerFontSize;
        }
      }

      // 2. Calculate Dynamic Indent (Respects padding-left)
      // Visual Indent = Distance from UL left edge to LI Content left edge.
      // PptxGenJS 'indent' = Space between bullet and text?
      // Actually PptxGenJS 'indent' allows setting the hanging indent.
      // We calculate the TOTAL visual offset from the parent container.
      // 1 px = 0.75 pt (approx, standard DTP).
      // We must scale it by config.scale.
      const visualIndentPx = liRect.left - parentRect.left;
      /*
         Standard indent in PPT is ~27pt.
         If visualIndentPx is small (e.g. 10px padding), we want small indent.
         If visualIndentPx is large (40px padding), we want large indent.
         We treat 'indent' as the value to pass to PptxGenJS.
      */
      const computedIndentPt = visualIndentPx * 0.75 * config.scale;

      if (bullet && computedIndentPt > 0) {
        bullet.indent = computedIndentPt;
        // Also support custom margin between bullet and text if provided in listConfig?
        // For now, computedIndentPt covers the visual placement.
      }

      // 3. Extract Text Parts
      const parts = collectTextParts(child, liStyle, config.scale);

      if (parts.length > 0) {
        parts.forEach((p) => {
          if (!p.options) p.options = {};
        });

        // A. Apply Bullet
        // Workaround: pptxgenjs bullets inherit the style of the text run they are attached to.
        // To support ::marker styles (color, size) that differ from the text, we create
        // a "dummy" text run at the start of the list item that carries the bullet configuration.
        if (bullet) {
          const firstPartInfo = parts[0].options;

          // Create a dummy run. We use a Zero Width Space to ensure it's rendered but invisible.
          // This "run" will hold the bullet and its specific color/size.
          const bulletRun = {
            text: '\u200B',
            options: {
              ...firstPartInfo, // Inherit base props (fontFace, etc.)
              color: bullet.color || firstPartInfo.color,
              fontSize: bullet.fontSize || firstPartInfo.fontSize,
              bullet: bullet,
            },
          };

          // Don't duplicate transparent or empty color from firstPart if bullet has one
          if (bullet.color) bulletRun.options.color = bullet.color;
          if (bullet.fontSize) bulletRun.options.fontSize = bullet.fontSize;

          // Prepend
          parts.unshift(bulletRun);
        }

        // B. Apply Spacing
        let ptBefore = 0;
        let ptAfter = 0;

        // A. Check Global Options (Expected in Points)
        if (globalOptions.listConfig?.spacing) {
          if (typeof globalOptions.listConfig.spacing.before === 'number') {
            ptBefore = globalOptions.listConfig.spacing.before;
          }
          if (typeof globalOptions.listConfig.spacing.after === 'number') {
            ptAfter = globalOptions.listConfig.spacing.after;
          }
        }
        // B. Fallback to CSS Margins (Convert px -> pt)
        else {
          const mt = parseFloat(liStyle.marginTop) || 0;
          const mb = parseFloat(liStyle.marginBottom) || 0;
          if (mt > 0) ptBefore = mt * 0.75 * config.scale;
          if (mb > 0) ptAfter = mb * 0.75 * config.scale;
        }

        if (ptBefore > 0) parts[0].options.paraSpaceBefore = ptBefore;
        if (ptAfter > 0) parts[0].options.paraSpaceAfter = ptAfter;

        if (index < liChildren.length - 1) {
          parts[parts.length - 1].options.breakLine = true;
        }

        listItems.push(...parts);
      }
    });

    if (listItems.length > 0) {
      // Add background if exists
      const bgColorObj = parseColor(style.backgroundColor);
      if (bgColorObj.hex && bgColorObj.opacity > 0) {
        items.push({
          type: 'shape',
          zIndex,
          domOrder,
          shapeType: 'rect',
          options: { x, y, w, h, fill: { color: bgColorObj.hex } },
        });
      }

      items.push({
        type: 'text',
        zIndex: zIndex + 1,
        domOrder,
        textParts: listItems,
        options: {
          x,
          y,
          w,
          h,
          align: 'left',
          valign: 'top',
          margin: 0,
          autoFit: true,
          wrap: !(style.whiteSpace === 'nowrap' || style.whiteSpace === 'pre'),
          vert: writingModeVert,
        },
      });

      return { items, stopRecursion: true };
    }
  }

  if (node.tagName === 'CANVAS') {
    const item = {
      type: 'image',
      zIndex,
      domOrder,
      options: { x, y, w, h, rotate: rotation, data: null },
    };

    const job = async () => {
      try {
        // Direct data extraction from the canvas element
        // This preserves the exact current state of the chart
        const dataUrl = node.toDataURL('image/png');

        // Basic validation
        if (dataUrl && dataUrl.length > 10) {
          item.options.data = dataUrl;
        } else {
          item.skip = true;
        }
      } catch (e) {
        // Tainted canvas (CORS issues) will throw here
        console.warn('Failed to capture canvas content:', e);
        item.skip = true;
      }
    };

    return { items: [item], job, stopRecursion: true };
  }

  // --- ASYNC JOB: SVG Tags ---
  if (node.nodeName.toUpperCase() === 'SVG') {
    const item = {
      type: 'image',
      zIndex,
      domOrder,
      options: { data: null, x, y, w, h, rotate: rotation },
    };

    const job = async () => {
      // Use svgToSvg for vector output (Convert to Shape in PowerPoint)
      // Use svgToPng for rasterized output (pixel perfect)
      const converter = globalOptions.svgAsVector ? svgToSvg : svgToPng;
      const processed = await converter(node);
      if (processed) item.options.data = processed;
      else item.skip = true;
    };

    return { items: [item], job, stopRecursion: true };
  }

  // --- ASYNC JOB: IMG Tags ---
  if (node.tagName === 'IMG') {
    let radii = {
      tl: parseFloat(style.borderTopLeftRadius) || 0,
      tr: parseFloat(style.borderTopRightRadius) || 0,
      br: parseFloat(style.borderBottomRightRadius) || 0,
      bl: parseFloat(style.borderBottomLeftRadius) || 0,
    };

    const hasAnyRadius = radii.tl > 0 || radii.tr > 0 || radii.br > 0 || radii.bl > 0;
    if (!hasAnyRadius) {
      const parent = node.parentElement;
      const parentStyle = window.getComputedStyle(parent);
      if (parentStyle.overflow !== 'visible') {
        const pRadii = {
          tl: parseFloat(parentStyle.borderTopLeftRadius) || 0,
          tr: parseFloat(parentStyle.borderTopRightRadius) || 0,
          br: parseFloat(parentStyle.borderBottomRightRadius) || 0,
          bl: parseFloat(parentStyle.borderBottomLeftRadius) || 0,
        };
        const pRect = parent.getBoundingClientRect();
        if (Math.abs(pRect.width - rect.width) < 5 && Math.abs(pRect.height - rect.height) < 5) {
          radii = pRadii;
        }
      }
    }

    const objectFit = style.objectFit || 'fill'; // default CSS behavior is fill
    const objectPosition = style.objectPosition || '50% 50%';

    const item = {
      type: 'image',
      zIndex,
      domOrder,
      options: { x, y, w, h, rotate: rotation, data: null },
    };

    const job = async () => {
      const processed = await getProcessedImage(
        node.src,
        widthPx,
        heightPx,
        radii,
        objectFit,
        objectPosition
      );
      if (processed) item.options.data = processed;
      else item.skip = true;
    };

    return { items: [item], job, stopRecursion: true };
  }

  // --- ASYNC JOB: Icons and Other Elements ---
  if (isIconElement(node)) {
    const item = {
      type: 'image',
      zIndex,
      domOrder,
      options: { x, y, w, h, rotate: rotation, data: null },
    };
    const job = async () => {
      const pngData = await elementToCanvasImage(node, widthPx, heightPx);
      if (pngData) item.options.data = pngData;
      else item.skip = true;
    };
    return { items: [item], job, stopRecursion: true };
  }

  // Radii logic
  const borderRadiusValue = parseFloat(style.borderRadius) || 0;
  const borderBottomLeftRadius = parseFloat(style.borderBottomLeftRadius) || 0;
  const borderBottomRightRadius = parseFloat(style.borderBottomRightRadius) || 0;
  const borderTopLeftRadius = parseFloat(style.borderTopLeftRadius) || 0;
  const borderTopRightRadius = parseFloat(style.borderTopRightRadius) || 0;

  const hasPartialBorderRadius =
    borderTopLeftRadius !== borderTopRightRadius ||
    borderTopLeftRadius !== borderBottomRightRadius ||
    borderTopLeftRadius !== borderBottomLeftRadius;

  // --- PRIORITY SVG: Solid Fill with Partial Border Radius (Vector Cone/Tab) ---
  // Fix for "missing cone": Prioritize SVG vector generation over Raster Canvas for simple shapes with partial radii.
  // This avoids html2canvas failures on empty divs.
  const tempBg = parseColor(style.backgroundColor);
  const isTxt = isTextContainer(node);

  // BUG FIX: Don't treat as a vector shape if it has content (like text or children).
  // This prevents containers like ".glass-box" from being treated as empty shapes and stopping recursion.
  const hasContent = node.textContent.trim().length > 0 || node.children.length > 0;

  if (hasPartialBorderRadius && tempBg.hex && !isTxt && !hasContent) {
    const shapeSvg = generateCustomShapeSVG(widthPx, heightPx, tempBg.hex, tempBg.opacity, {
      tl: parseFloat(style.borderTopLeftRadius) || 0,
      tr: parseFloat(style.borderTopRightRadius) || 0,
      br: parseFloat(style.borderBottomRightRadius) || 0,
      bl: parseFloat(style.borderBottomLeftRadius) || 0,
    });

    return {
      items: [
        {
          type: 'image',
          zIndex,
          domOrder,
          options: { data: shapeSvg, x, y, w, h, rotate: rotation },
        },
      ],
      stopRecursion: true, // Treat as leaf
    };
  }

  // --- ASYNC JOB: Clipped Divs via Canvas ---
  // Only capture as image if it's an empty leaf.
  // Rasterizing containers (like .glass-box) kills editability of children.
  if (hasPartialBorderRadius && isClippedByParent(node) && !hasContent) {
    const marginLeft = parseFloat(style.marginLeft) || 0;
    const marginTop = parseFloat(style.marginTop) || 0;
    x += marginLeft * PX_TO_INCH * config.scale;
    y += marginTop * PX_TO_INCH * config.scale;

    const item = {
      type: 'image',
      zIndex,
      domOrder,
      options: { x, y, w, h, rotate: rotation, data: null },
    };

    const job = async () => {
      const canvasImageData = await elementToCanvasImage(node, widthPx, heightPx);
      if (canvasImageData) item.options.data = canvasImageData;
      else item.skip = true;
    };

    return { items: [item], job, stopRecursion: true };
  }

  // --- SYNC: Standard CSS Extraction ---
  const bgColorObj = parseColor(style.backgroundColor);
  const bgClip = style.webkitBackgroundClip || style.backgroundClip;
  const isBgClipText = bgClip === 'text';
  const bgImgStr = style.backgroundImage;
  const hasGradient = !isBgClipText && bgImgStr && bgImgStr.includes('linear-gradient');
  const urlMatch =
    !isBgClipText && !hasGradient && bgImgStr ? bgImgStr.match(/url\(['"]?(.*?)['"]?\)/) : null;
  const hasBgImgUrl = !!urlMatch;

  const borderColorObj = parseColor(style.borderColor);
  const borderWidth = parseFloat(style.borderWidth);
  const hasBorder = borderWidth > 0 && borderColorObj.hex;

  const borderInfo = getBorderInfo(style, config.scale);
  const hasUniformBorder = borderInfo.type === 'uniform';
  const hasCompositeBorder = borderInfo.type === 'composite';

  const shadowStr = style.boxShadow;
  const hasShadow = shadowStr && shadowStr !== 'none';
  const softEdge = getSoftEdges(style.filter, config.scale);

  let isImageWrapper = false;
  const imgChild = Array.from(node.children).find((c) => c.tagName === 'IMG');
  if (imgChild) {
    const childW = imgChild.offsetWidth || imgChild.getBoundingClientRect().width;
    const childH = imgChild.offsetHeight || imgChild.getBoundingClientRect().height;
    if (childW >= widthPx - 2 && childH >= heightPx - 2) isImageWrapper = true;
  }

  let textPayload = null;
  const isText = isTextContainer(node);

  if (isText) {
    const textParts = collectTextParts(node, style, config.scale, null, true, inheritedOpacity);

    if (textParts.length > 0) {
      let align = style.textAlign || 'left';
      if (align === 'start') align = 'left';
      if (align === 'end') align = 'right';
      let valign = 'top';
      if (style.verticalAlign === 'middle') valign = 'middle';
      if (style.verticalAlign === 'bottom') valign = 'bottom';

      // Fix: Handle Flexbox axis swap for vertical writing modes OR column-direction flex
      const isVertical = writingModeVert && writingModeVert !== 'none';
      const isColumn = style.flexDirection === 'column' || style.flexDirection === 'column-reverse';

      if (isVertical || isColumn) {
        // Vertical Axis Swap (Main axis is vertical)
        if (style.alignItems === 'center') align = 'center';
        if (style.alignItems === 'flex-end' || style.alignItems === 'end') align = 'right';

        if (style.justifyContent === 'center' && style.display.includes('flex')) valign = 'middle';
        if (style.justifyContent === 'flex-end' && style.display.includes('flex'))
          valign = 'bottom';
      } else {
        // Standard Row Axis (Main axis is horizontal)
        if (style.alignItems === 'center') valign = 'middle';
        if (style.alignItems === 'flex-end' || style.alignItems === 'end') valign = 'bottom';

        if (style.justifyContent === 'center' && style.display.includes('flex')) align = 'center';
        if (style.justifyContent === 'flex-end' || style.justifyContent === 'end') {
          if (style.display.includes('flex')) align = 'right';
        }
      }

      // Fix: Suppress lineSpacing for vertical text to prevent improper character gaps
      if (isVertical) {
        textParts.forEach((p) => {
          if (p.options) delete p.options.lineSpacing;
        });
      }

      let padding = getPadding(style, config.scale);

      textPayload = { text: textParts, align, valign, inset: padding };
    }
  }

  let bgJob = null;

  if (hasBgImgUrl || hasGradient || (softEdge && bgColorObj.hex && !isImageWrapper)) {
    if (hasBgImgUrl) {
      const bgUrl = urlMatch[1];
      const radii = {
        tl: parseFloat(style.borderTopLeftRadius) || 0,
        tr: parseFloat(style.borderTopRightRadius) || 0,
        br: parseFloat(style.borderBottomRightRadius) || 0,
        bl: parseFloat(style.borderBottomLeftRadius) || 0,
      };

      const bgItem = {
        type: 'image',
        zIndex,
        domOrder,
        options: { x, y, w, h, rotate: rotation, data: null },
      };
      items.push(bgItem);

      bgJob = async () => {
        const processed = await getProcessedImage(
          bgUrl,
          widthPx,
          heightPx,
          radii,
          style.backgroundSize || 'cover',
          style.backgroundPosition || '50% 50%'
        );
        if (processed) bgItem.options.data = processed;
        else bgItem.skip = true;
      };
    } else {
      let bgData = null;
      let padIn = 0;
      if (softEdge) {
        const svgInfo = generateBlurredSVG(
          widthPx,
          heightPx,
          bgColorObj.hex,
          borderRadiusValue,
          softEdge
        );
        bgData = svgInfo.data;
        padIn = svgInfo.padding * PX_TO_INCH * config.scale;
      } else {
        bgData = generateGradientSVG(
          widthPx,
          heightPx,
          style.backgroundImage,
          hasPartialBorderRadius
            ? {
                tl: borderTopLeftRadius,
                tr: borderTopRightRadius,
                br: borderBottomRightRadius,
                bl: borderBottomLeftRadius,
              }
            : borderRadiusValue,
          hasBorder ? { color: borderColorObj.hex, width: borderWidth } : null
        );
      }

      if (bgData) {
        items.push({
          type: 'image',
          zIndex,
          domOrder,
          options: {
            data: bgData,
            x: x - padIn,
            y: y - padIn,
            w: w + padIn * 2,
            h: h + padIn * 2,
            rotate: rotation,
          },
        });
      }
    }

    if (textPayload) {
      const fs0 = textPayload.text[0]?.options?.fontSize;
      textPayload.text[0].options.fontSize =
        typeof fs0 === 'number' ? Math.floor(fs0 * 10) / 10 : 12;
      items.push({
        type: 'text',
        zIndex: zIndex + 1,
        domOrder,
        textParts: textPayload.text,
        options: {
          x,
          y,
          w: w * 1.05, // Safety buffer to prevent wrapping
          h,
          align: textPayload.align,
          valign: textPayload.valign,
          inset: textPayload.inset,
          rotate: rotation,
          margin: 0,
          wrap: !(style.whiteSpace === 'nowrap' || style.whiteSpace === 'pre'),
          autoFit: true,
          vert: writingModeVert,
        },
      });
    }
    if (hasCompositeBorder) {
      const borderItems = createCompositeBorderItems(
        borderInfo.sides,
        x,
        y,
        w,
        h,
        config.scale,
        zIndex,
        domOrder
      );
      items.push(...borderItems);
    }
  } else if (
    (bgColorObj.hex && !isImageWrapper) ||
    hasUniformBorder ||
    hasCompositeBorder ||
    hasShadow ||
    textPayload
  ) {
    const finalAlpha = safeOpacity * bgColorObj.opacity;
    const transparency = (1 - finalAlpha) * 100;
    const useSolidFill = bgColorObj.hex && !isImageWrapper;

    if (hasPartialBorderRadius && useSolidFill && !textPayload) {
      const shapeSvg = generateCustomShapeSVG(
        widthPx,
        heightPx,
        bgColorObj.hex,
        bgColorObj.opacity,
        {
          tl: parseFloat(style.borderTopLeftRadius) || 0,
          tr: parseFloat(style.borderTopRightRadius) || 0,
          br: parseFloat(style.borderBottomRightRadius) || 0,
          bl: parseFloat(style.borderBottomLeftRadius) || 0,
        }
      );

      items.push({
        type: 'image',
        zIndex,
        domOrder,
        options: { data: shapeSvg, x, y, w, h, rotate: rotation },
      });
    } else {
      const shapeOpts = {
        x,
        y,
        w,
        h,
        rotate: rotation,
        fill: useSolidFill
          ? { color: bgColorObj.hex, transparency: transparency }
          : { type: 'none' },
        line: hasUniformBorder ? borderInfo.options : null,
      };

      if (hasShadow) shapeOpts.shadow = getVisibleShadow(shadowStr, config.scale);

      // 1. Calculate dimensions first
      const minDimension = Math.min(widthPx, heightPx);

      let rawRadius = parseFloat(style.borderRadius) || 0;
      const isPercentage = style.borderRadius && style.borderRadius.toString().includes('%');

      // 2. Normalize radius to pixels
      let radiusPx = rawRadius;
      if (isPercentage) {
        radiusPx = (rawRadius / 100) * minDimension;
      }

      let shapeType = pptx.ShapeType.rect;

      // 3. Determine Shape Logic
      const isSquare = Math.abs(widthPx - heightPx) < 1;
      const isFullyRound = radiusPx >= minDimension / 2;

      // CASE A: It is an Ellipse if:
      // 1. It is explicitly "50%" (standard CSS way to make ovals/circles)
      // 2. OR it is a perfect square and fully rounded (a circle)
      if (isFullyRound && (isPercentage || isSquare)) {
        shapeType = pptx.ShapeType.ellipse;
      }
      // CASE B: It is a Rounded Rectangle (including "Pill" shapes)
      else if (radiusPx > 0) {
        shapeType = pptx.ShapeType.roundRect;
        let cappedRadiusPx = Math.min(radiusPx, minDimension / 2);
        shapeOpts.rectRadius = cappedRadiusPx * PX_TO_INCH * config.scale;
      }

      if (textPayload) {
        const fs0 = textPayload.text[0]?.options?.fontSize;
        textPayload.text[0].options.fontSize =
          typeof fs0 === 'number' ? Math.floor(fs0 * 10) / 10 : 12;
        const textOptions = {
          shape: shapeType,
          ...shapeOpts,
          w: w * 1.015, // Safety buffer to prevent wrapping
          h,
          rotate: rotation,
          align: textPayload.align,
          valign: textPayload.valign,
          inset: textPayload.inset,
          margin: 0,
          wrap: !(style.whiteSpace === 'nowrap' || style.whiteSpace === 'pre'),
          autoFit: true,
          vert: writingModeVert,
        };
        items.push({
          type: 'text',
          zIndex,
          domOrder,
          textParts: textPayload.text,
          options: textOptions,
        });
      } else if (!hasPartialBorderRadius) {
        items.push({
          type: 'shape',
          zIndex,
          domOrder,
          shapeType,
          options: shapeOpts,
        });
      }
    }

    if (hasCompositeBorder) {
      const borderSvgData = generateCompositeBorderSVG(
        widthPx,
        heightPx,
        borderRadiusValue,
        borderInfo.sides
      );
      if (borderSvgData) {
        items.push({
          type: 'image',
          zIndex: zIndex + 1,
          domOrder,
          options: { data: borderSvgData, x, y, w, h, rotate: rotation },
        });
      }
    }
  }

  return { items, job: bgJob, stopRecursion: !!textPayload };
}

function isComplexHierarchy(root) {
  // Use a simple tree traversal to find forbidden elements in the list structure
  const stack = [root];
  while (stack.length > 0) {
    const el = stack.pop();

    // 1. Layouts: Flex/Grid on LIs
    if (el.tagName === 'LI') {
      const s = window.getComputedStyle(el);
      if (s.display === 'flex' || s.display === 'grid' || s.display === 'inline-flex') return true;
    }

    // 2. Media / Icons
    if (['IMG', 'SVG', 'CANVAS', 'VIDEO', 'IFRAME'].includes(el.tagName)) return true;
    if (isIconElement(el)) return true;

    // 3. Nested Lists (Flattening logic doesn't support nested bullets well yet)
    if (el !== root && (el.tagName === 'UL' || el.tagName === 'OL')) return true;

    // Recurse, but don't go too deep if not needed
    for (let i = 0; i < el.children.length; i++) {
      stack.push(el.children[i]);
    }
  }
  return false;
}

function createCompositeBorderItems(sides, x, y, w, h, scale, zIndex, domOrder) {
  const items = [];
  const pxToInch = 1 / 96;
  const common = { zIndex: zIndex + 1, domOrder, shapeType: 'rect' };

  if (sides.top.width > 0)
    items.push({
      ...common,
      options: { x, y, w, h: sides.top.width * pxToInch * scale, fill: { color: sides.top.color } },
    });
  if (sides.right.width > 0)
    items.push({
      ...common,
      options: {
        x: x + w - sides.right.width * pxToInch * scale,
        y,
        w: sides.right.width * pxToInch * scale,
        h,
        fill: { color: sides.right.color },
      },
    });
  if (sides.bottom.width > 0)
    items.push({
      ...common,
      options: {
        x,
        y: y + h - sides.bottom.width * pxToInch * scale,
        w,
        h: sides.bottom.width * pxToInch * scale,
        fill: { color: sides.bottom.color },
      },
    });
  if (sides.left.width > 0)
    items.push({
      ...common,
      options: {
        x,
        y,
        w: sides.left.width * pxToInch * scale,
        h,
        fill: { color: sides.left.color },
      },
    });

  return items;
}
