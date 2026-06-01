// src/__tests__/pptx-normalizer.test.js
import { describe, it, expect } from 'vitest';
import JSZip from 'jszip';
import { normalizePptxZip } from '../pptx-normalizer.js';

const CONTENT_TYPES_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';

function buildContentTypes({ defaults = [], overrides = [] } = {}) {
  const defaultsXml = defaults
    .map((d) => `  <Default Extension="${d.ext}" ContentType="${d.contentType}"/>`)
    .join('\n');
  const overridesXml = overrides
    .map((o) => `  <Override PartName="${o.partName}" ContentType="${o.contentType}"/>`)
    .join('\n');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="${CONTENT_TYPES_NS}">
${defaultsXml}
${overridesXml}
</Types>`;
}

describe('normalizePptxZip', () => {
  it('removes Override entries that point at files missing from the zip', async () => {
    const zip = new JSZip();
    zip.file(
      '[Content_Types].xml',
      buildContentTypes({
        defaults: [
          { ext: 'rels', contentType: 'application/vnd.openxmlformats-package.relationships+xml' },
          { ext: 'xml', contentType: 'application/xml' },
        ],
        overrides: [
          {
            partName: '/ppt/slideMasters/slideMaster1.xml',
            contentType:
              'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml',
          },
          {
            partName: '/ppt/slideMasters/slideMaster2.xml',
            contentType:
              'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml',
          },
          {
            partName: '/ppt/slideMasters/slideMaster3.xml',
            contentType:
              'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml',
          },
        ],
      })
    );
    zip.file('ppt/slideMasters/slideMaster1.xml', '<sldMaster/>');

    await normalizePptxZip(zip);

    const xml = await zip.file('[Content_Types].xml').async('string');
    expect(xml).toContain('slideMaster1.xml');
    expect(xml).not.toContain('slideMaster2.xml');
    expect(xml).not.toContain('slideMaster3.xml');
  });

  it('preserves Default entries even when their extension has no parts in the zip', async () => {
    const zip = new JSZip();
    zip.file(
      '[Content_Types].xml',
      buildContentTypes({
        defaults: [{ ext: 'fntdata', contentType: 'application/x-fontdata' }],
        overrides: [],
      })
    );

    await normalizePptxZip(zip);

    const xml = await zip.file('[Content_Types].xml').async('string');
    expect(xml).toContain('Extension="fntdata"');
  });

  it('is idempotent', async () => {
    const zip = new JSZip();
    zip.file(
      '[Content_Types].xml',
      buildContentTypes({
        defaults: [{ ext: 'xml', contentType: 'application/xml' }],
        overrides: [
          {
            partName: '/ppt/slideMasters/slideMaster1.xml',
            contentType:
              'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml',
          },
          {
            partName: '/ppt/slideMasters/slideMaster9.xml',
            contentType:
              'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml',
          },
        ],
      })
    );
    zip.file('ppt/slideMasters/slideMaster1.xml', '<sldMaster/>');

    await normalizePptxZip(zip);
    const firstPass = await zip.file('[Content_Types].xml').async('string');
    await normalizePptxZip(zip);
    const secondPass = await zip.file('[Content_Types].xml').async('string');

    expect(secondPass).toBe(firstPass);
  });

  it('does nothing when [Content_Types].xml is missing', async () => {
    const zip = new JSZip();
    await expect(normalizePptxZip(zip)).resolves.toBeUndefined();
  });

  it('skips a malformed [Content_Types].xml without throwing', async () => {
    const zip = new JSZip();
    const garbage = '<<<not xml>>>';
    zip.file('[Content_Types].xml', garbage);

    await expect(normalizePptxZip(zip)).resolves.toBeUndefined();
    const xml = await zip.file('[Content_Types].xml').async('string');
    expect(xml).toBe(garbage);
  });

  it('handles PartName values without a leading slash', async () => {
    const zip = new JSZip();
    zip.file(
      '[Content_Types].xml',
      buildContentTypes({
        overrides: [
          {
            partName: 'ppt/slideMasters/slideMaster1.xml',
            contentType:
              'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml',
          },
          {
            partName: 'ppt/slideMasters/slideMasterGhost.xml',
            contentType:
              'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml',
          },
        ],
      })
    );
    zip.file('ppt/slideMasters/slideMaster1.xml', '<sldMaster/>');

    await normalizePptxZip(zip);

    const xml = await zip.file('[Content_Types].xml').async('string');
    expect(xml).toContain('slideMaster1.xml');
    expect(xml).not.toContain('slideMasterGhost.xml');
  });

  it('merges multiple a:pPr elements and sorts children according to DrawingML schema', async () => {
    const zip = new JSZip();
    zip.file(
      '[Content_Types].xml',
      buildContentTypes({
        defaults: [{ ext: 'xml', contentType: 'application/xml' }],
      })
    );
    const slideXml = `<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:sp>
        <p:txBody>
          <a:p>
            <a:pPr algn="ctr">
              <a:spcAft><a:spcPts val="600"/></a:spcAft>
            </a:pPr>
            <a:pPr indent="300">
              <a:lnSpc><a:spcPct val="120000"/></a:lnSpc>
            </a:pPr>
            <a:r>
              <a:t>Hello</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>`;
    zip.file('ppt/slides/slide1.xml', slideXml);

    await normalizePptxZip(zip);

    const normalizedXml = await zip.file('ppt/slides/slide1.xml').async('string');
    const doc = new DOMParser().parseFromString(normalizedXml, 'text/xml');
    const pPrs = doc.getElementsByTagName('a:pPr');
    expect(pPrs.length).toBe(1);

    const pPr = pPrs[0];
    expect(pPr.getAttribute('algn')).toBe('ctr');
    expect(pPr.getAttribute('indent')).toBe('300');

    const children = Array.from(pPr.childNodes).filter((n) => n.nodeType === 1);
    expect(children.length).toBe(2);
    expect(children[0].localName).toBe('lnSpc');
    expect(children[1].localName).toBe('spcAft');
  });

  it('restores character spacing spc attribute from typeface __spc_ suffix', async () => {
    const zip = new JSZip();
    zip.file(
      '[Content_Types].xml',
      buildContentTypes({
        defaults: [{ ext: 'xml', contentType: 'application/xml' }],
      })
    );
    const slideXml = `<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:sp>
        <p:txBody>
          <a:p>
            <a:r>
              <a:rPr>
                <a:latin typeface="Arial__spc_150"/>
              </a:rPr>
              <a:t>Hello</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>`;
    zip.file('ppt/slides/slide1.xml', slideXml);

    await normalizePptxZip(zip);

    const normalizedXml = await zip.file('ppt/slides/slide1.xml').async('string');
    const doc = new DOMParser().parseFromString(normalizedXml, 'text/xml');
    const latin = doc.getElementsByTagName('a:latin')[0];
    expect(latin.getAttribute('typeface')).toBe('Arial');

    const rPr = doc.getElementsByTagName('a:rPr')[0];
    expect(rPr.getAttribute('spc')).toBe('150');
  });

  it('sorts visual elements in spTree based on __z_ altText prefix and removes transport prefix', async () => {
    const zip = new JSZip();
    zip.file(
      '[Content_Types].xml',
      buildContentTypes({
        defaults: [{ ext: 'xml', contentType: 'application/xml' }],
      })
    );
    const slideXml = `<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr/>
      <p:grpSpPr/>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="Shape 2" descr="__z_1__dom_1 Original User Alt"/>
        </p:nvSpPr>
      </p:sp>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="1" name="Shape 1" descr="__z_0__dom_0"/>
        </p:nvSpPr>
      </p:sp>
      <p:extLst/>
    </p:spTree>
  </p:cSld>
</p:sld>`;
    zip.file('ppt/slides/slide1.xml', slideXml);

    await normalizePptxZip(zip);

    const normalizedXml = await zip.file('ppt/slides/slide1.xml').async('string');
    const doc = new DOMParser().parseFromString(normalizedXml, 'text/xml');
    const spTree = doc.getElementsByTagName('p:spTree')[0];
    const elements = Array.from(spTree.childNodes).filter((n) => n.nodeType === 1);

    expect(elements.length).toBe(5);
    expect(elements[0].localName).toBe('nvGrpSpPr');
    expect(elements[1].localName).toBe('grpSpPr');

    // Shape 1 (id="1") should be sorted first among visual elements because z=0 < z=1
    expect(elements[2].getElementsByTagName('p:cNvPr')[0].getAttribute('id')).toBe('1');
    expect(elements[2].getElementsByTagName('p:cNvPr')[0].hasAttribute('descr')).toBe(false);

    // Shape 2 (id="2") should be sorted second among visual elements because z=1
    expect(elements[3].getElementsByTagName('p:cNvPr')[0].getAttribute('id')).toBe('2');
    expect(elements[3].getElementsByTagName('p:cNvPr')[0].getAttribute('descr')).toBe(
      'Original User Alt'
    );

    expect(elements[4].localName).toBe('extLst');
  });

  it('sorts visual elements in spTree based on __z_ name attribute prefix and removes transport prefix', async () => {
    const zip = new JSZip();
    zip.file(
      '[Content_Types].xml',
      buildContentTypes({
        defaults: [{ ext: 'xml', contentType: 'application/xml' }],
      })
    );
    const slideXml = `<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr/>
      <p:grpSpPr/>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="__z_1__dom_1 Original User Name"/>
        </p:nvSpPr>
      </p:sp>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="1" name="__z_0__dom_0"/>
        </p:nvSpPr>
      </p:sp>
      <p:extLst/>
    </p:spTree>
  </p:cSld>
</p:sld>`;
    zip.file('ppt/slides/slide1.xml', slideXml);

    await normalizePptxZip(zip);

    const normalizedXml = await zip.file('ppt/slides/slide1.xml').async('string');
    const doc = new DOMParser().parseFromString(normalizedXml, 'text/xml');
    const spTree = doc.getElementsByTagName('p:spTree')[0];
    const elements = Array.from(spTree.childNodes).filter((n) => n.nodeType === 1);

    expect(elements.length).toBe(5);
    expect(elements[0].localName).toBe('nvGrpSpPr');
    expect(elements[1].localName).toBe('grpSpPr');

    // Shape 1 (id="1") should be sorted first among visual elements because z=0 < z=1
    expect(elements[2].getElementsByTagName('p:cNvPr')[0].getAttribute('id')).toBe('1');
    expect(elements[2].getElementsByTagName('p:cNvPr')[0].getAttribute('name')).toBe('Object 0');

    // Shape 2 (id="2") should be sorted second among visual elements because z=1
    expect(elements[3].getElementsByTagName('p:cNvPr')[0].getAttribute('id')).toBe('2');
    expect(elements[3].getElementsByTagName('p:cNvPr')[0].getAttribute('name')).toBe(
      'Original User Name'
    );

    expect(elements[4].localName).toBe('extLst');
  });

  it.skip('sorts child elements of presentation.xml according to PresentationML schema', async () => {
    const zip = new JSZip();
    zip.file(
      '[Content_Types].xml',
      buildContentTypes({
        defaults: [{ ext: 'xml', contentType: 'application/xml' }],
      })
    );
    const presentationXml = `<?xml version="1.0" encoding="UTF-8"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldSz cx="12192000" cy="6858000"/>
  <p:notesMasterIdLst/>
  <p:sldMasterIdLst/>
  <p:sldIdLst/>
</p:presentation>`;
    zip.file('ppt/presentation.xml', presentationXml);

    await normalizePptxZip(zip);

    const normalizedXml = await zip.file('ppt/presentation.xml').async('string');
    const doc = new DOMParser().parseFromString(normalizedXml, 'text/xml');
    const presentation = doc.getElementsByTagName('p:presentation')[0];
    const children = Array.from(presentation.childNodes).filter((n) => n.nodeType === 1);

    expect(children.length).toBe(4);
    expect(children[0].localName).toBe('sldMasterIdLst');
    expect(children[1].localName).toBe('notesMasterIdLst');
    expect(children[2].localName).toBe('sldIdLst');
    expect(children[3].localName).toBe('sldSz');
  });
});
