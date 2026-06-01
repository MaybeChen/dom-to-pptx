// src/pptx-normalizer.js
//
// Defensive OOXML normalizer that runs over the PPTX produced by PptxGenJS
// before we hand the .pptx blob to the user. Microsoft PowerPoint refuses to
// open files when [Content_Types].xml advertises parts that are not actually
// present in the package — see 错误诊断.md for the original incident report.
//
// This module operates on an already-loaded JSZip instance and mutates it in
// place. The caller is responsible for re-serializing the zip with DEFLATE
// compression afterwards.

const pPrOrder = [
  'lnSpc',
  'spcBef',
  'spcAft',
  'buClrTx',
  'buClr',
  'buSzTx',
  'buSzPct',
  'buSzPts',
  'buFontTx',
  'buFont',
  'buNone',
  'buAutoNum',
  'buChar',
  'buBlip',
  'tabLst',
  'defRPr',
  'extLst',
];

const presentationOrder = [
  'sldMasterIdLst',
  'notesMasterIdLst',
  'handoutMasterIdLst',
  'sldIdLst',
  'sldSz',
  'notesSz',
  'smartTags',
  'embeddedFontLst',
  'customShowLst',
  'photoAlbum',
  'custDataLst',
  'kins',
  'defaultTextStyle',
  'modifyVerifier',
  'extLst',
];

/**
 * Strips dangling <Override> entries from [Content_Types].xml.
 *
 * An Override is "dangling" when its PartName attribute references a file path
 * that does not exist inside the zip. Default entries are left untouched
 * because they apply to every file with a matching extension, and removing
 * them would break legitimate parts (e.g. the fntdata default added by the
 * font embedder).
 *
 * The function is idempotent: running it twice on the same zip yields the
 * same result as running it once.
 *
 * @param {import('jszip')} zip - JSZip instance with the loaded PPTX package.
 * @returns {Promise<void>}
 */
export async function normalizePptxZip(zip) {
  if (!zip) return;

  const contentTypesFile = zip.file('[Content_Types].xml');
  if (!contentTypesFile) return;

  let xmlStr;
  try {
    xmlStr = await contentTypesFile.async('string');
  } catch (e) {
    console.warn('[pptx-normalizer] Failed to read [Content_Types].xml:', e);
    return;
  }

  let doc;
  try {
    doc = new DOMParser().parseFromString(xmlStr, 'text/xml');
  } catch (e) {
    console.warn('[pptx-normalizer] Failed to parse [Content_Types].xml:', e);
    return;
  }

  const parserError = doc.getElementsByTagName('parsererror')[0];
  if (parserError) {
    console.warn('[pptx-normalizer] [Content_Types].xml has parser errors, skipping cleanup.');
    return;
  }

  const overrides = Array.from(doc.getElementsByTagName('Override'));
  let removedCount = 0;

  for (const node of overrides) {
    const partName = node.getAttribute('PartName');
    if (!partName) continue;

    // PartName is always an absolute path inside the zip, e.g. "/ppt/slideMasters/slideMaster2.xml".
    // JSZip indexes its files without the leading slash.
    const zipPath = partName.startsWith('/') ? partName.slice(1) : partName;

    if (!zip.file(zipPath)) {
      node.parentNode?.removeChild(node);
      removedCount++;
    }
  }

  if (removedCount > 0) {
    const serialized = serializeXmlWithDeclaration(doc);
    zip.file('[Content_Types].xml', serialized);
  }

  // Process all XML files inside the ppt/ directory
  for (const [relativePath, file] of Object.entries(zip.files)) {
    if (relativePath.startsWith('ppt/') && relativePath.endsWith('.xml')) {
      let xmlStr = await file.async('string');
      let doc;
      try {
        doc = new DOMParser().parseFromString(xmlStr, 'text/xml');
      } catch (e) {
        console.warn(`[pptx-normalizer] Failed to parse ${relativePath}:`, e);
        continue;
      }

      const parserError = doc.getElementsByTagName('parsererror')[0];
      if (parserError) {
        console.warn(`[pptx-normalizer] ${relativePath} has parser errors, skipping.`);
        continue;
      }

      let mutated = false;

      if (cleanParagraphProperties(doc)) {
        mutated = true;
      }

      if (restoreCharSpacing(doc)) {
        mutated = true;
      }

      if (relativePath.startsWith('ppt/slides/slide')) {
        if (sortSpTree(doc)) {
          mutated = true;
        }
      }

      if (mutated) {
        const serialized = serializeXmlWithDeclaration(doc);
        zip.file(relativePath, serialized);
      }
    }
  }
}

function cleanParagraphProperties(doc) {
  let mutated = false;
  const elements = Array.from(doc.getElementsByTagName('*'));
  const paragraphs = elements.filter((n) => n.localName === 'p');

  for (const p of paragraphs) {
    const pPrs = Array.from(p.childNodes).filter(
      (node) => node.nodeType === 1 && node.localName === 'pPr'
    );
    if (pPrs.length > 0) {
      const targetPPr = pPrs[0];

      // Merge subsequent pPr elements
      if (pPrs.length > 1) {
        for (let i = 1; i < pPrs.length; i++) {
          const sourcePPr = pPrs[i];
          // merge attributes
          for (const attr of Array.from(sourcePPr.attributes)) {
            if (targetPPr.getAttribute(attr.name) !== attr.value) {
              targetPPr.setAttribute(attr.name, attr.value);
              mutated = true;
            }
          }
          // merge children
          while (sourcePPr.firstChild) {
            targetPPr.appendChild(sourcePPr.firstChild);
            mutated = true;
          }
          p.removeChild(sourcePPr);
          mutated = true;
        }
      }

      // Resolve duplicate child elements in targetPPr (keep the last one of each localName)
      const childElements = Array.from(targetPPr.childNodes).filter((node) => node.nodeType === 1);
      const seen = new Map();
      for (let i = childElements.length - 1; i >= 0; i--) {
        const child = childElements[i];
        const key = child.localName;
        if (seen.has(key)) {
          targetPPr.removeChild(child);
          mutated = true;
        } else {
          seen.set(key, child);
        }
      }

      // Sort children of targetPPr
      const finalChildren = Array.from(targetPPr.childNodes).filter((node) => node.nodeType === 1);
      const sortedChildren = [...finalChildren].sort((a, b) => {
        const idxA = pPrOrder.indexOf(a.localName);
        const idxB = pPrOrder.indexOf(b.localName);
        const orderA = idxA !== -1 ? idxA : 999;
        const orderB = idxB !== -1 ? idxB : 999;
        return orderA - orderB;
      });

      let needsReorder = false;
      for (let i = 0; i < finalChildren.length; i++) {
        if (finalChildren[i] !== sortedChildren[i]) {
          needsReorder = true;
          break;
        }
      }

      if (needsReorder) {
        while (targetPPr.firstChild) {
          targetPPr.removeChild(targetPPr.firstChild);
        }
        for (const child of sortedChildren) {
          targetPPr.appendChild(child);
        }
        mutated = true;
      }

      // Ensure targetPPr is the first child element of p
      const firstElementChild = Array.from(p.childNodes).find((node) => node.nodeType === 1);
      if (firstElementChild && firstElementChild !== targetPPr) {
        p.insertBefore(targetPPr, firstElementChild);
        mutated = true;
      }
    }
  }
  return mutated;
}

function restoreCharSpacing(doc) {
  let mutated = false;
  const elements = Array.from(doc.getElementsByTagName('*'));
  const typefaceElements = elements.filter(
    (n) => n.hasAttribute('typeface') && n.getAttribute('typeface').includes('__spc_')
  );

  for (const el of typefaceElements) {
    const typeface = el.getAttribute('typeface');
    const parts = typeface.split('__spc_');
    el.setAttribute('typeface', parts[0]);
    mutated = true;

    // Find the parent rPr or defRPr or endParaRPr to set 'spc'
    let parent = el.parentNode;
    if (
      parent &&
      (parent.localName === 'rPr' ||
        parent.localName === 'defRPr' ||
        parent.localName === 'endParaRPr')
    ) {
      const spcVal = parts[1];
      if (parent.getAttribute('spc') !== spcVal) {
        parent.setAttribute('spc', spcVal);
      }
    }
  }
  return mutated;
}

function sortSpTree(doc) {
  let mutated = false;
  const elements = Array.from(doc.getElementsByTagName('*'));
  const spTrees = elements.filter((n) => n.localName === 'spTree');

  for (const spTree of spTrees) {
    const childNodes = Array.from(spTree.childNodes);
    const elementChildren = childNodes.filter((node) => node.nodeType === 1);

    const firstElements = [];
    const lastElements = [];
    const visualElements = [];

    for (const el of elementChildren) {
      const localName = el.localName;
      if (localName === 'nvGrpSpPr' || localName === 'grpSpPr') {
        firstElements.push(el);
      } else if (localName === 'extLst') {
        lastElements.push(el);
      } else {
        visualElements.push(el);
      }
    }

    // Parse z-order for visualElements
    const elementInfos = visualElements.map((el, originalIndex) => {
      let zVal = Infinity;
      let domVal = originalIndex;

      // Find cNvPr element
      let cNvPr = null;
      const nvPr = Array.from(el.childNodes).find(
        (n) => n.nodeType === 1 && n.localName.startsWith('nv')
      );
      if (nvPr) {
        cNvPr = Array.from(nvPr.childNodes).find(
          (n) => n.nodeType === 1 && n.localName === 'cNvPr'
        );
      }
      if (!cNvPr) {
        cNvPr = Array.from(el.getElementsByTagName('*')).find((n) => n.localName === 'cNvPr');
      }

      if (cNvPr) {
        const descr = cNvPr.getAttribute('descr') || '';
        const nameAttr = cNvPr.getAttribute('name') || '';
        let hasVal = false;

        const descrMatch = descr.match(/^__z_(\d+)__dom_(\d+)(.*)/);
        if (descrMatch) {
          zVal = parseInt(descrMatch[1], 10);
          domVal = parseInt(descrMatch[2], 10);
          hasVal = true;
          const userText = descrMatch[3].trim();
          if (userText) {
            cNvPr.setAttribute('descr', userText);
          } else {
            cNvPr.removeAttribute('descr');
          }
          mutated = true;
        }

        const nameMatch = nameAttr.match(/^__z_(\d+)__dom_(\d+)(.*)/);
        if (nameMatch) {
          if (!hasVal) {
            zVal = parseInt(nameMatch[1], 10);
            domVal = parseInt(nameMatch[2], 10);
            hasVal = true;
          }
          const userText = nameMatch[3].trim();
          if (userText) {
            cNvPr.setAttribute('name', userText);
          } else {
            cNvPr.setAttribute('name', `Object ${domVal}`);
          }
          mutated = true;
        }
      }
      return { el, zVal, domVal, originalIndex };
    });

    // Sort visualElements
    elementInfos.sort((a, b) => {
      if (a.zVal !== b.zVal) {
        return a.zVal - b.zVal;
      }
      return a.domVal - b.domVal;
    });

    const sortedVisualElements = elementInfos.map((info) => info.el);

    // Check if the order has changed
    let orderChanged = false;
    for (let i = 0; i < visualElements.length; i++) {
      if (visualElements[i] !== sortedVisualElements[i]) {
        orderChanged = true;
        break;
      }
    }

    if (orderChanged) {
      // Re-construct spTree children
      while (spTree.firstChild) {
        spTree.removeChild(spTree.firstChild);
      }

      for (const el of firstElements) {
        spTree.appendChild(el);
      }
      for (const el of sortedVisualElements) {
        spTree.appendChild(el);
      }
      for (const el of lastElements) {
        spTree.appendChild(el);
      }
      mutated = true;
    }
  }
  return mutated;
}

function serializeXmlWithDeclaration(doc) {
  let serialized = new XMLSerializer().serializeToString(doc);
  if (!serialized.startsWith('<?xml')) {
    serialized = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + serialized;
  }
  return serialized;
}
