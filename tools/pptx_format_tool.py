from ibm_watsonx_orchestrate.agent_builder.tools import tool, WXOFile
import zipfile, io, json, xml.etree.ElementTree as ET

# ─── Namespaces ───────────────────────────────────────────────────────────────

NS = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}

# ─── Unit Converters ──────────────────────────────────────────────────────────

def emu_to_pt(emu):
    """EMU → points (1 pt = 12700 EMU)"""
    try:
        return round(int(emu) / 12700, 2)
    except:
        return emu

def emu_to_in(emu):
    """EMU → inches (1 inch = 914400 EMU)"""
    try:
        return round(int(emu) / 914400, 4)
    except:
        return emu

# ─── Color Parsing ────────────────────────────────────────────────────────────

def parse_color(elem):
    """
    Handles all OOXML color types:
      srgbClr  → explicit hex  e.g. #FF0000
      sysClr   → system color  (returns lastClr hex if available)
      schemeClr→ theme slot    e.g. 'dk1', 'accent1'
      prstClr  → preset name   e.g. 'black'
      scrgbClr → percentage RGB
    """
    if elem is None:
        return None
    tag = elem.tag.split('}')[-1]

    if tag == 'srgbClr':
        val = elem.get('val', '')
        mods = _color_mods(elem)
        return {'type': 'rgb', 'hex': f'#{val.upper()}', **mods}

    elif tag == 'sysClr':
        last = elem.get('lastClr', '')
        return {'type': 'system', 'val': elem.get('val'),
                'hex': f'#{last.upper()}' if last else None}

    elif tag == 'schemeClr':
        mods = _color_mods(elem)
        return {'type': 'scheme', 'val': elem.get('val'), **mods}

    elif tag == 'prstClr':
        return {'type': 'preset', 'val': elem.get('val')}

    elif tag == 'scrgbClr':
        r = round(int(elem.get('r', 0)) / 100000 * 255)
        g = round(int(elem.get('g', 0)) / 100000 * 255)
        b = round(int(elem.get('b', 0)) / 100000 * 255)
        return {'type': 'rgb', 'hex': f'#{r:02X}{g:02X}{b:02X}'}

    return None

def _color_mods(elem):
    """Extract lumMod, lumOff, tint, shade, alpha modifiers from a color element."""
    mods = {}
    for mod in ['lumMod', 'lumOff', 'tint', 'shade', 'alpha']:
        m = elem.find(f'a:{mod}', NS)
        if m is not None:
            mods[mod] = round(int(m.get('val', 0)) / 1000, 1)
    return mods

def get_first_color(parent):
    """Return the first color child of any fill/line element."""
    for tag in ['srgbClr', 'sysClr', 'schemeClr', 'prstClr', 'scrgbClr']:
        child = parent.find(f'a:{tag}', NS)
        if child is not None:
            return parse_color(child)
    return None

# ─── Fill Parsing ─────────────────────────────────────────────────────────────

def parse_fill(elem):
    """
    Detects and parses whichever fill type is present:
      noFill    → no fill
      solidFill → single color
      gradFill  → gradient (multiple color stops + angle)
      pattFill  → pattern
      blipFill  → image/picture fill
    """
    if elem.find('a:noFill', NS) is not None:
        return {'type': 'none'}

    sf = elem.find('a:solidFill', NS)
    if sf is not None:
        return {'type': 'solid', 'color': get_first_color(sf)}

    gf = elem.find('a:gradFill', NS)
    if gf is not None:
        stops = []
        for gs in gf.findall('.//a:gs', NS):
            stops.append({
                'position_pct': round(int(gs.get('pos', 0)) / 1000, 1),
                'color': get_first_color(gs)
            })
        lin = gf.find('a:lin', NS)
        angle = None
        if lin is not None and lin.get('ang'):
            angle = round(int(lin.get('ang')) / 60000, 1)
        return {'type': 'gradient', 'stops': stops, 'angle_deg': angle}

    pf = elem.find('a:pattFill', NS)
    if pf is not None:
        return {'type': 'pattern', 'preset': pf.get('prst')}

    if elem.find('a:blipFill', NS) is not None or elem.find('p:blipFill', NS) is not None:
        return {'type': 'image'}

    return None

# ─── Line / Border Parsing ────────────────────────────────────────────────────

def parse_line(spPr):
    """
    Extracts border/outline from a shape's spPr:
      width, dash style, compound type, cap type, color, arrows
    """
    ln = spPr.find('a:ln', NS)
    if ln is None:
        return None

    result = {}
    if ln.get('w'):
        result['width_pt'] = emu_to_pt(ln.get('w'))
    if ln.get('cap'):
        result['cap'] = ln.get('cap')
    if ln.get('cmpd'):
        result['compound'] = ln.get('cmpd')

    prstDash = ln.find('a:prstDash', NS)
    if prstDash is not None:
        result['dash'] = prstDash.get('val')  # solid, dash, dot, dashDot, etc.

    if ln.find('a:noFill', NS) is not None:
        result['color'] = None  # transparent border
    else:
        sf = ln.find('a:solidFill', NS)
        if sf is not None:
            result['color'] = get_first_color(sf)

    for arrow, key in [('a:headEnd', 'arrowHead'), ('a:tailEnd', 'arrowTail')]:
        end = ln.find(arrow, NS)
        if end is not None:
            result[key] = {
                'type': end.get('type'),   # none, arrow, triangle, stealth, diamond, oval
                'width': end.get('w'),
                'length': end.get('len')
            }

    return result or None

# ─── Effects Parsing ──────────────────────────────────────────────────────────

def parse_effects(elem):
    """
    Parses effectLst children:
      outerShdw → drop shadow (blur, distance, direction, color)
      innerShdw → inner shadow
      glow      → glow radius + color
      reflection → boolean flag
    """
    effectLst = elem.find('a:effectLst', NS)
    if effectLst is None:
        return None

    effects = {}

    outerShdw = effectLst.find('a:outerShdw', NS)
    if outerShdw is not None:
        effects['outerShadow'] = {
            'blurRadius_pt': emu_to_pt(outerShdw.get('blurRad', 0)),
            'distance_pt':   emu_to_pt(outerShdw.get('dist', 0)),
            'direction_deg': round(int(outerShdw.get('dir', 0)) / 60000, 1),
            'alignment':     outerShdw.get('algn'),
            'color':         get_first_color(outerShdw)
        }

    innerShdw = effectLst.find('a:innerShdw', NS)
    if innerShdw is not None:
        effects['innerShadow'] = {
            'blurRadius_pt': emu_to_pt(innerShdw.get('blurRad', 0)),
            'distance_pt':   emu_to_pt(innerShdw.get('dist', 0)),
            'color':         get_first_color(innerShdw)
        }

    glow = effectLst.find('a:glow', NS)
    if glow is not None:
        effects['glow'] = {
            'radius_pt': emu_to_pt(glow.get('rad', 0)),
            'color':     get_first_color(glow)
        }

    if effectLst.find('a:reflection', NS) is not None:
        effects['reflection'] = True

    return effects or None

# ─── Text Run Parsing ─────────────────────────────────────────────────────────

def parse_run_props(rPr):
    """
    Full text run formatting:
      font family, size, bold, italic, underline, strikethrough,
      color, highlight, kerning, spacing, baseline, language, effects
    """
    props = {}

    if rPr.get('sz'):
        props['fontSize_pt'] = round(int(rPr.get('sz')) / 100, 2)
    if rPr.get('b') is not None:
        props['bold'] = rPr.get('b') == '1'
    if rPr.get('i') is not None:
        props['italic'] = rPr.get('i') == '1'
    if rPr.get('u'):
        props['underline'] = rPr.get('u')   # sng, dbl, heavy, dotted, etc.
    if rPr.get('strike'):
        props['strikethrough'] = rPr.get('strike')
    if rPr.get('baseline'):
        props['baseline'] = int(rPr.get('baseline'))  # positive=superscript, negative=subscript
    if rPr.get('kern'):
        props['kerning_pt'] = round(int(rPr.get('kern')) / 100, 2)
    if rPr.get('spc'):
        props['charSpacing_pt'] = round(int(rPr.get('spc')) / 100, 2)
    if rPr.get('lang'):
        props['language'] = rPr.get('lang')

    # Font faces
    latin = rPr.find('a:latin', NS)
    if latin is not None:
        props['font'] = latin.get('typeface')

    ea = rPr.find('a:ea', NS)
    if ea is not None:
        props['eastAsianFont'] = ea.get('typeface')

    # Text color
    sf = rPr.find('a:solidFill', NS)
    if sf is not None:
        props['color'] = get_first_color(sf)

    # Highlight color
    hl = rPr.find('a:highlight', NS)
    if hl is not None:
        props['highlight'] = get_first_color(hl)

    # Text-level effects (e.g. text shadow)
    fx = parse_effects(rPr)
    if fx:
        props['effects'] = fx

    return props

def parse_paragraph(para):
    """
    Paragraph-level formatting + individual text runs.
    Covers alignment, indent, bullets, line spacing, space before/after.
    """
    result = {}

    pPr = para.find('a:pPr', NS)
    if pPr is not None:
        pp = {}

        if pPr.get('algn'):
            pp['alignment'] = pPr.get('algn')   # l, ctr, r, just, dist
        if pPr.get('indent'):
            pp['indent_pt'] = emu_to_pt(pPr.get('indent'))
        if pPr.get('marL'):
            pp['marginLeft_pt'] = emu_to_pt(pPr.get('marL'))
        if pPr.get('marR'):
            pp['marginRight_pt'] = emu_to_pt(pPr.get('marR'))
        if pPr.get('lvl'):
            pp['level'] = int(pPr.get('lvl'))

        # Space before paragraph
        spcBef = pPr.find('a:spcBef', NS)
        if spcBef is not None:
            spcPts = spcBef.find('a:spcPts', NS)
            spcPct = spcBef.find('a:spcPct', NS)
            if spcPts is not None:
                pp['spaceBefore_pt'] = round(int(spcPts.get('val', 0)) / 100, 2)
            if spcPct is not None:
                pp['spaceBefore_pct'] = round(int(spcPct.get('val', 0)) / 1000, 1)

        # Space after paragraph
        spcAft = pPr.find('a:spcAft', NS)
        if spcAft is not None:
            spcPts = spcAft.find('a:spcPts', NS)
            spcPct = spcAft.find('a:spcPct', NS)
            if spcPts is not None:
                pp['spaceAfter_pt'] = round(int(spcPts.get('val', 0)) / 100, 2)
            if spcPct is not None:
                pp['spaceAfter_pct'] = round(int(spcPct.get('val', 0)) / 1000, 1)

        # Line spacing
        lnSpc = pPr.find('a:lnSpc', NS)
        if lnSpc is not None:
            spcPts = lnSpc.find('a:spcPts', NS)
            spcPct = lnSpc.find('a:spcPct', NS)
            if spcPts is not None:
                pp['lineSpacing_pt'] = round(int(spcPts.get('val', 0)) / 100, 2)
            if spcPct is not None:
                pp['lineSpacing_pct'] = round(int(spcPct.get('val', 0)) / 1000, 1)

        # Bullet character
        buChar = pPr.find('a:buChar', NS)
        if buChar is not None:
            pp['bulletChar'] = buChar.get('char')
        buNone = pPr.find('a:buNone', NS)
        if buNone is not None:
            pp['bulletNone'] = True
        buAutoNum = pPr.find('a:buAutoNum', NS)
        if buAutoNum is not None:
            pp['bulletAutoNum'] = buAutoNum.get('type')
        buFont = pPr.find('a:buFont', NS)
        if buFont is not None:
            pp['bulletFont'] = buFont.get('typeface')
        buClr = pPr.find('a:buClr', NS)
        if buClr is not None:
            pp['bulletColor'] = get_first_color(buClr)
        buSzPct = pPr.find('a:buSzPct', NS)
        if buSzPct is not None:
            pp['bulletSize_pct'] = round(int(buSzPct.get('val', 0)) / 1000, 1)

        # Default run props on the paragraph
        defRPr = pPr.find('a:defRPr', NS)
        if defRPr is not None:
            pp['defaultRunProps'] = parse_run_props(defRPr)

        if pp:
            result['paragraphProps'] = pp

    # Individual text runs
    runs = []
    for child in para:
        tag = child.tag.split('}')[-1]
        if tag == 'r':
            t = child.find('a:t', NS)
            run = {'text': t.text if t is not None else ''}
            rPr = child.find('a:rPr', NS)
            if rPr is not None:
                run['formatting'] = parse_run_props(rPr)
            runs.append(run)
        elif tag == 'br':
            runs.append({'type': 'lineBreak'})

    if runs:
        result['runs'] = runs

    return result

# ─── Shape Parsing ────────────────────────────────────────────────────────────

def parse_shape(sp):
    """
    Parses a single <p:sp> element:
      - identity (id, name)
      - position + size in inches
      - geometry (preset shape name or custom)
      - fill, border/line, effects
      - text body with all paragraphs
    """
    shape = {'type': 'shape'}

    # Identity
    cNvPr = sp.find('.//p:cNvPr', NS)
    if cNvPr is not None:
        shape['id'] = cNvPr.get('id')
        shape['name'] = cNvPr.get('name')

    # Shape properties
    spPr = sp.find('p:spPr', NS)
    if spPr is not None:

        # Position & size
        xfrm = spPr.find('a:xfrm', NS)
        if xfrm is not None:
            off = xfrm.find('a:off', NS)
            ext = xfrm.find('a:ext', NS)
            if off is not None:
                shape['position'] = {
                    'x_in': emu_to_in(off.get('x', 0)),
                    'y_in': emu_to_in(off.get('y', 0))
                }
            if ext is not None:
                shape['size'] = {
                    'width_in':  emu_to_in(ext.get('cx', 0)),
                    'height_in': emu_to_in(ext.get('cy', 0))
                }
            if xfrm.get('rot'):
                shape['rotation_deg'] = round(int(xfrm.get('rot')) / 60000, 2)
            if xfrm.get('flipH') == '1':
                shape['flipH'] = True
            if xfrm.get('flipV') == '1':
                shape['flipV'] = True

        # Geometry
        prstGeom = spPr.find('a:prstGeom', NS)
        if prstGeom is not None:
            shape['geometry'] = prstGeom.get('prst')  # e.g. 'rect', 'roundRect', 'ellipse'
        elif spPr.find('a:custGeom', NS) is not None:
            shape['geometry'] = 'custom'

        # Fill
        fill = parse_fill(spPr)
        if fill:
            shape['fill'] = fill

        # Border / outline
        line = parse_line(spPr)
        if line:
            shape['line'] = line

        # Effects
        effects = parse_effects(spPr)
        if effects:
            shape['effects'] = effects

    # Text body
    txBody = sp.find('p:txBody', NS)
    if txBody is not None:
        tb = {}

        bodyPr = txBody.find('a:bodyPr', NS)
        if bodyPr is not None:
            bpr = {}
            if bodyPr.get('vert'):
                bpr['verticalText'] = bodyPr.get('vert')
            if bodyPr.get('anchor'):
                bpr['verticalAnchor'] = bodyPr.get('anchor')  # t, ctr, b
            if bodyPr.get('wrap'):
                bpr['wrap'] = bodyPr.get('wrap')
            # Internal margins
            for m in ['lIns', 'rIns', 'tIns', 'bIns']:
                if bodyPr.get(m):
                    bpr[f'{m}_pt'] = emu_to_pt(bodyPr.get(m))
            # Autofit
            if bodyPr.find('a:normAutofit', NS) is not None:
                bpr['autofit'] = 'normal'
            elif bodyPr.find('a:spAutoFit', NS) is not None:
                bpr['autofit'] = 'shape'
            elif bodyPr.find('a:noAutofit', NS) is not None:
                bpr['autofit'] = 'none'
            if bpr:
                tb['bodyProperties'] = bpr

        tb['paragraphs'] = [parse_paragraph(p) for p in txBody.findall('a:p', NS)]
        shape['textBody'] = tb

    return shape

def parse_connector(cxnSp):
    """Connector lines between shapes — same structure as a shape."""
    shape = parse_shape(cxnSp)
    shape['type'] = 'connector'
    return shape

def parse_picture(pic):
    """Image/picture element — position, size, crop geometry, border."""
    shape = {'type': 'picture'}

    cNvPr = pic.find('.//p:cNvPr', NS)
    if cNvPr is not None:
        shape['id'] = cNvPr.get('id')
        shape['name'] = cNvPr.get('name')

    spPr = pic.find('p:spPr', NS)
    if spPr is not None:
        xfrm = spPr.find('a:xfrm', NS)
        if xfrm is not None:
            off = xfrm.find('a:off', NS)
            ext = xfrm.find('a:ext', NS)
            if off is not None:
                shape['position'] = {'x_in': emu_to_in(off.get('x', 0)),
                                     'y_in': emu_to_in(off.get('y', 0))}
            if ext is not None:
                shape['size'] = {'width_in': emu_to_in(ext.get('cx', 0)),
                                 'height_in': emu_to_in(ext.get('cy', 0))}
        prstGeom = spPr.find('a:prstGeom', NS)
        if prstGeom is not None:
            shape['cropGeometry'] = prstGeom.get('prst')
        line = parse_line(spPr)
        if line:
            shape['line'] = line

    shape['fill'] = {'type': 'image'}
    return shape

def parse_table(tbl):
    """
    Full table: style flags, column widths, row heights,
    per-cell fills, all 6 border sides, text content.
    """
    table = {'columns': [], 'rows': [], 'style': {}}

    tblPr = tbl.find('a:tblPr', NS)
    if tblPr is not None:
        st = {}
        for flag in ['bandRow', 'bandCol', 'firstRow', 'lastRow', 'firstCol', 'lastCol']:
            if tblPr.get(flag) == '1':
                st[flag] = True
        fill = parse_fill(tblPr)
        if fill:
            st['fill'] = fill
        if st:
            table['style'] = st

    # Column widths
    tblGrid = tbl.find('a:tblGrid', NS)
    if tblGrid is not None:
        for col in tblGrid.findall('a:gridCol', NS):
            table['columns'].append({'width_in': emu_to_in(col.get('w', 0))})

    # Rows and cells
    for tr in tbl.findall('a:tr', NS):
        row = {'height_in': emu_to_in(tr.get('h', 0)), 'cells': []}
        for tc in tr.findall('a:tc', NS):
            cell = {}

            # Merge flags
            if tc.get('gridSpan'):
                cell['colSpan'] = int(tc.get('gridSpan'))
            if tc.get('rowSpan'):
                cell['rowSpan'] = int(tc.get('rowSpan'))
            if tc.get('hMerge') == '1':
                cell['hMerge'] = True
            if tc.get('vMerge') == '1':
                cell['vMerge'] = True

            tcPr = tc.find('a:tcPr', NS)
            if tcPr is not None:
                cp = {}
                for m in ['marL', 'marR', 'marT', 'marB']:
                    if tcPr.get(m):
                        cp[f'{m}_pt'] = emu_to_pt(tcPr.get(m))
                if tcPr.get('anchor'):
                    cp['verticalAnchor'] = tcPr.get('anchor')
                fill = parse_fill(tcPr)
                if fill:
                    cp['fill'] = fill

                # All 6 possible cell borders
                for border in ['lnL', 'lnR', 'lnT', 'lnB', 'lnTlToBr', 'lnBlToTr']:
                    be = tcPr.find(f'a:{border}', NS)
                    if be is not None:
                        ln = parse_line(be) if be.tag.split('}')[-1] == 'ln' else None
                        # the border elements ARE the ln elements here
                        w = be.get('w')
                        line_data = {}
                        if w:
                            line_data['width_pt'] = emu_to_pt(w)
                        sf = be.find('a:solidFill', NS)
                        if sf is not None:
                            line_data['color'] = get_first_color(sf)
                        if be.find('a:noFill', NS) is not None:
                            line_data = None
                        if line_data:
                            cp[f'border_{border}'] = line_data

                if cp:
                    cell['properties'] = cp

            # Cell text
            txBody = tc.find('a:txBody', NS)
            if txBody is not None:
                cell['paragraphs'] = [parse_paragraph(p) for p in txBody.findall('a:p', NS)]

            row['cells'].append(cell)
        table['rows'].append(row)

    return table

def parse_group(grpSp):
    """Group shapes — recurses into children."""
    group = {'type': 'group', 'children': []}

    cNvPr = grpSp.find('.//p:cNvPr', NS)
    if cNvPr is not None:
        group['id'] = cNvPr.get('id')
        group['name'] = cNvPr.get('name')

    grpSpPr = grpSp.find('p:grpSpPr', NS)
    if grpSpPr is not None:
        xfrm = grpSpPr.find('a:xfrm', NS)
        if xfrm is not None:
            off = xfrm.find('a:off', NS)
            ext = xfrm.find('a:ext', NS)
            if off is not None:
                group['position'] = {'x_in': emu_to_in(off.get('x', 0)),
                                     'y_in': emu_to_in(off.get('y', 0))}
            if ext is not None:
                group['size'] = {'width_in': emu_to_in(ext.get('cx', 0)),
                                 'height_in': emu_to_in(ext.get('cy', 0))}

    for child in grpSp:
        tag = child.tag.split('}')[-1]
        if tag == 'sp':
            group['children'].append(parse_shape(child))
        elif tag == 'pic':
            group['children'].append(parse_picture(child))
        elif tag == 'cxnSp':
            group['children'].append(parse_connector(child))
        elif tag == 'grpSp':
            group['children'].append(parse_group(child))  # recursive

    return group

# ─── Slide Entry Point ────────────────────────────────────────────────────────

def parse_slide(tree, _NS=None):
    """
    Top-level parser for a single slide XML tree.
    Returns a dict with:
      - background fill
      - list of all shapes (shapes, pictures, connectors, groups, tables)
    """
    root = tree.getroot()
    slide = {'background': None, 'shapes': []}

    # Slide background
    bg = root.find('.//p:bg', NS)
    if bg is not None:
        bgPr = bg.find('p:bgPr', NS)
        if bgPr is not None:
            slide['background'] = parse_fill(bgPr)

    # Shape tree
    spTree = root.find('.//p:spTree', NS)
    if spTree is None:
        return slide

    for elem in spTree:
        tag = elem.tag.split('}')[-1]

        if tag == 'sp':
            slide['shapes'].append(parse_shape(elem))

        elif tag == 'pic':
            slide['shapes'].append(parse_picture(elem))

        elif tag == 'cxnSp':
            slide['shapes'].append(parse_connector(elem))

        elif tag == 'grpSp':
            slide['shapes'].append(parse_group(elem))

        elif tag == 'graphicFrame':
            # Tables and charts live inside graphicFrames
            gf = {'type': 'graphicFrame'}
            cNvPr = elem.find('.//p:cNvPr', NS)
            if cNvPr is not None:
                gf['id'] = cNvPr.get('id')
                gf['name'] = cNvPr.get('name')

            xfrm = elem.find('p:xfrm', NS)
            if xfrm is not None:
                off = xfrm.find('a:off', NS)
                ext = xfrm.find('a:ext', NS)
                if off is not None:
                    gf['position'] = {'x_in': emu_to_in(off.get('x', 0)),
                                      'y_in': emu_to_in(off.get('y', 0))}
                if ext is not None:
                    gf['size'] = {'width_in': emu_to_in(ext.get('cx', 0)),
                                  'height_in': emu_to_in(ext.get('cy', 0))}

            tbl = elem.find('.//a:tbl', NS)
            if tbl is not None:
                gf['contentType'] = 'table'
                gf['table'] = parse_table(tbl)
            else:
                gf['contentType'] = 'chart'  # chart XML is in a separate embedded file

            slide['shapes'].append(gf)

    return slide

# ─── Watsonx Orchestrate Tool ─────────────────────────────────────────────────

@tool(
    name="extract_pptx_formatting",
    description="Extracts complete formatting from a PowerPoint file and returns it as a downloadable JSON file."
)
def extract_pptx_formatting(pptx_file: WXOFile) -> bytes:
    """Extracts all visual formatting from a .pptx and returns a JSON file.

    Args:
        pptx_file (WXOFile): The uploaded PowerPoint file.

    Returns:
        bytes: A JSON file containing full formatting details per slide.
    """
    file_bytes = WXOFile.get_content(pptx_file)

    result = {}
    with zipfile.ZipFile(io.BytesIO(file_bytes)) as z:
        slide_files = sorted(
            [f for f in z.namelist()
             if f.startswith('ppt/slides/slide') and f.endswith('.xml')],
            key=lambda x: int(''.join(filter(str.isdigit, x)) or 0)
        )
        for slide_file in slide_files:
            with z.open(slide_file) as f:
                tree = ET.parse(f)
                result[slide_file] = parse_slide(tree)

    return json.dumps(result, ensure_ascii=False, indent=2).encode('utf-8')