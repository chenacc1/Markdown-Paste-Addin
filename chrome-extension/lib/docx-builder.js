/**
 * docx-builder.js — Self-contained browser-side .docx generator
 *
 * Generates Microsoft Word documents directly in the browser extension
 * without requiring a Python bridge server. Supports:
 *   - Headings (H1-H6)
 *   - Paragraphs with inline bold/italic/code
 *   - Tables with header row styling
 *   - Code blocks (monospace, gray background)
 *   - Task lists (☐/☒)
 *   - Blockquotes
 *   - Horizontal rules
 *   - Images (URL → embedded)
 *   - Math formulas (LaTeX → OMML native Word equations)
 *
 * Uses openxml-js (embedded) for XML building and a minimal OPC/ZIP packer.
 */

const DocxBuilder = (() => {
  'use strict';

  // ─── Namespace constants ─────────────────────────────────────────────────────
  const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const NS_CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
  const NS_M = 'http://schemas.openxmlformats.org/officeDocument/2006/math';

  // ─── XML helpers ──────────────────────────────────────────────────────────────

  function esc(str) {
    return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;')
      .replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;');
  }

  function el(name, attrs, content) {
    const a = [];
    if (attrs) {
      for (const k of Object.keys(attrs)) {
        if (attrs[k] != null) a.push(`${k}="${esc(attrs[k])}"`);
      }
    }
    const aStr = a.length ? ' ' + a.join(' ') : '';
    if (content == null) return `<w:${name}${aStr}/>`;
    return `<w:${name}${aStr}>${content}</w:${name}>`;
  }

  function relEl(name, attrs, content) {
    const a = [];
    if (attrs) {
      for (const k of Object.keys(attrs)) {
        if (attrs[k] != null) a.push(`${k}="${esc(attrs[k])}"`);
      }
    }
    const aStr = a.length ? ' ' + a.join(' ') : '';
    if (content == null) return `<${name}${aStr}/>`;
    return `<${name}${aStr}>${content}</${name}>`;
  }

  // ─── Run-level helpers ────────────────────────────────────────────────────────

  function runProps(font, sz, bold, italic, color) {
    const parts = [];
    parts.push(el('rFonts', { 'w:ascii': font, 'w:hAnsi': font, 'w:eastAsia': font }));
    if (sz) parts.push(el('sz', { 'w:val': String(Math.round(sz * 2)) }));
    if (bold) parts.push(el('b'));
    if (italic) parts.push(el('i'));
    if (color) parts.push(el('color', { 'w:val': color }));
    return el('rPr', null, parts.join(''));
  }

  function textRun(text, font, sz, bold, italic, color) {
    const rPr = runProps(font || 'Times New Roman', sz || 12, bold, italic, color);
    return el('r', null, rPr + el('t', { 'xml:space': 'preserve' }, esc(text)));
  }

  function emptyRun() {
    return el('r', null, el('rPr', null, el('rFonts', { 'w:ascii': 'Times New Roman', 'w:eastAsia': '宋体' })));
  }

  // ─── Paragraph helpers ────────────────────────────────────────────────────────

  function paraProps(align, spacingBefore, spacingAfter, lineSpacing, indent) {
    const parts = [];
    if (align) parts.push(el('jc', { 'w:val': align }));
    if (spacingBefore) parts.push(el('spacing', { 'w:before': String(spacingBefore), 'w:after': String(spacingAfter || 0), 'w:line': String(lineSpacing || 360), 'w:lineRule': 'auto' }));
    else if (lineSpacing) parts.push(el('spacing', { 'w:line': String(lineSpacing), 'w:lineRule': 'auto' }));
    if (indent) parts.push(el('ind', { 'w:left': String(indent), 'w:right': String(Math.round(indent * 0.3)) }));
    return el('pPr', null, parts.join(''));
  }

  function shadedParaProps(align, spacingBefore, spacingAfter, shadingColor, indent) {
    const shd = el('shd', { 'w:val': 'clear', 'w:color': 'auto', 'w:fill': shadingColor });
    const pBdr = el('pBdr', null, el('bottom', { 'w:val': 'single', 'w:sz': '4', 'w:space': '1', 'w:color': shadingColor }));
    let p = '';
    if (align) p += el('jc', { 'w:val': align });
    p += shd + pBdr;
    if (indent) p += el('ind', { 'w:left': String(indent) });
    return el('pPr', null, p);
  }

  // ─── Table helpers ────────────────────────────────────────────────────────────

  function tableBorders() {
    const borders = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'];
    return el('tblBorders', null, borders.map(b => el(b, { 'w:val': 'single', 'w:sz': '4', 'w:space': '0', 'w:color': 'auto' })).join(''));
  }

  function cellShading(color) {
    return el('shd', { 'w:val': 'clear', 'w:color': 'auto', 'w:fill': color });
  }

  // ─── Inline markdown processor ────────────────────────────────────────────────

  function parseInlineMarkdown(text) {
    const tokens = [];
    const pattern = /(\*\*\*(.+?)\*\*\*|___(.+?)___|\*\*(.+?)\*\*|__(.+?)__|(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)|(?<!_)_(?!_)(.+?)(?<!_)_(?!_)|`([^`]+)`)/g;
    let lastIdx = 0;
    let m;
    while ((m = pattern.exec(text)) !== null) {
      if (m.index > lastIdx) {
        tokens.push({ type: 'plain', text: text.substring(lastIdx, m.index) });
      }
      if (m[2] || m[3]) {
        tokens.push({ type: 'bold_italic', text: m[2] || m[3] });
      } else if (m[4] || m[5]) {
        tokens.push({ type: 'bold', text: m[4] || m[5] });
      } else if (m[6] || m[7]) {
        tokens.push({ type: 'italic', text: m[6] || m[7] });
      } else if (m[8]) {
        tokens.push({ type: 'code', text: m[8] });
      }
      lastIdx = m.index + m[0].length;
    }
    if (lastIdx < text.length) {
      tokens.push({ type: 'plain', text: text.substring(lastIdx) });
    }
    if (tokens.length === 0) tokens.push({ type: 'plain', text: text });
    return tokens;
  }

  function runsFromText(text, font, sz) {
    const tokens = parseInlineMarkdown(String(text));
    const runs = [];
    for (const tok of tokens) {
      switch (tok.type) {
        case 'bold':
          runs.push(textRun(tok.text, font, sz, true, false));
          break;
        case 'italic':
          runs.push(textRun(tok.text, font, sz, false, true));
          break;
        case 'bold_italic':
          runs.push(textRun(tok.text, font, sz, true, true));
          break;
        case 'code':
          runs.push(textRun(tok.text, 'Consolas', 9.5, false, false, 'B43C3C'));
          break;
        default:
          runs.push(textRun(tok.text, font, sz, false, false));
      }
    }
    return runs;
  }

  // ═══════════════════════════════════════════════════════════════════════════════
  // LaTeX → OMML converter (ported from parser_math.py)
  // ═══════════════════════════════════════════════════════════════════════════════

  const GREEK = {
    alpha:'α',beta:'β',gamma:'γ',delta:'δ',epsilon:'ε',varepsilon:'ε',
    zeta:'ζ',eta:'η',theta:'θ',vartheta:'ϑ',iota:'ι',kappa:'κ',
    lambda:'λ',mu:'μ',nu:'ν',xi:'ξ',pi:'π',varpi:'ϖ',rho:'ρ',
    varrho:'ϱ',sigma:'σ',varsigma:'ς',tau:'τ',upsilon:'υ',phi:'φ',
    varphi:'ϕ',chi:'χ',psi:'ψ',omega:'ω',
    Gamma:'Γ',Delta:'Δ',Theta:'Θ',Lambda:'Λ',Xi:'Ξ',Pi:'Π',
    Sigma:'Σ',Upsilon:'Υ',Phi:'Φ',Psi:'Ψ',Omega:'Ω',
  };

  const SYMBOLS = {
    pm:'±',mp:'∓',times:'×',div:'÷',cdot:'·',star:'⋆',
    leq:'≤',le:'≤',geq:'≥',ge:'≥',neq:'≠',ne:'≠',
    approx:'≈',equiv:'≡',sim:'∼',simeq:'≃',propto:'∝',
    ll:'≪',gg:'≫',prec:'≺',succ:'≻',
    infty:'∞',partial:'∂',nabla:'∇',
    forall:'∀',exists:'∃',neg:'¬',
    to:'→',rightarrow:'→',leftarrow:'←',Rightarrow:'⇒',Leftarrow:'⇐',
    leftrightarrow:'↔',Leftrightarrow:'⇔',
    uparrow:'↑',downarrow:'↓',Uparrow:'⇑',Downarrow:'⇓',
    mapsto:'↦',hookrightarrow:'↪',
    ldots:'…',cdots:'⋯',vdots:'⋮',ddots:'⋱',
    quad:'  ',qquad:'    ',
    colon:':',backslash:'\\',
    langle:'⟨',rangle:'⟩',lfloor:'⌊',rfloor:'⌋',lceil:'⌈',rceil:'⌉',
    Vert:'‖','||':'‖',
    aleph:'ℵ',hbar:'ℏ',
    land:'∧',lor:'∨',
    'in':'∈',notin:'∉',ni:'∋',
    subset:'⊂',supset:'⊃',subseteq:'⊆',supseteq:'⊇',
    cup:'∪',cap:'∩',setminus:'\\',emptyset:'∅',
    perp:'⊥',angle:'∠',
    triangle:'△',square:'□',
    dag:'†',ddag:'‡',S:'§',P:'¶',
    copyright:'©',textregistered:'®',
    degree:'°',circ:'∘',
    prime:"'",backprime:'‵',
    ell:'ℓ',wp:'℘',Re:'ℜ',Im:'ℑ',
  };

  const NARY_OPS = {
    sum:'∑',prod:'∏',coprod:'∐',
    int:'∫',oint:'∮',iint:'∬',iiint:'∭',
    bigcup:'⋃',bigcap:'⋂',bigoplus:'⨁',bigotimes:'⨂',
    bigvee:'⋁',bigwedge:'⋀',
  };

  const ACCENT_CHARS = {
    bar:'0304',overline:'0304',vec:'20D7',overrightarrow:'20D7',
    hat:'0302',widehat:'0302',tilde:'0303',widetilde:'0303',
    dot:'0307',ddot:'0308',breve:'0306',check:'030C',
    acute:'0301',grave:'0300',
  };

  const DELIM_MAP = {
    '(':'(', ')':')', '[':'[', ']':']',
    '{':'{', '}':'}', '|':'|', '||':'‖',
    '.':'', '/':'/',
    langle:'⟨', rangle:'⟩', lfloor:'⌊', rfloor:'⌋', lceil:'⌈', rceil:'⌉',
    uparrow:'↑', downarrow:'↓', Uparrow:'⇑', Downarrow:'⇓',
  };

  const FONT_STY = {
    mathbf:'b', mathit:'i', mathsf:'sans', mathtt:'monospace',
    mathrm:'roman', mathbb:'double-struck', mathcal:'script',
    mathscr:'script', mathfrak:'fraktur',
  };

  // ─── LaTeX parser ─────────────────────────────────────────────────────────────

  class LaTeXParser {
    constructor(src) {
      this.src = src;
      this.pos = 0;
      this._maxIter = 50000;
    }

    peek() { return this.pos < this.src.length ? this.src[this.pos] : ''; }
    consume() { return this.pos < this.src.length ? this.src[this.pos++] : ''; }
    skipWs() { while (this.pos < this.src.length && /\s/.test(this.src[this.pos])) this.pos++; }

    parse() {
      const items = this._parseItems();
      return { type: 'root', items };
    }

    _parseItems(stopChars = '') {
      const items = [];
      let iter = 0;
      while (this.pos < this.src.length && iter++ < this._maxIter) {
        const ch = this.peek();
        if (stopChars.includes(ch)) break;
        if (ch === '}') break;

        if (ch === '\\') {
          items.push(this._parseCommand());
        } else if (ch === '_') {
          this.consume();
          items.push({ type: 'sub', content: this._parseSingle() });
        } else if (ch === '^') {
          this.consume();
          items.push({ type: 'sup', content: this._parseSingle() });
        } else if (ch === '{') {
          this.consume();
          const inner = this._parseItems('}');
          if (this.peek() === '}') this.consume();
          items.push(...inner);
        } else if (/\s/.test(ch)) {
          this.consume();
        } else {
          items.push({ type: 'text', text: this.consume() });
        }
      }
      return items;
    }

    _parseSingle() {
      this.skipWs();
      const ch = this.peek();
      if (ch === '{') {
        this.consume();
        const items = this._parseItems('}');
        if (this.peek() === '}') this.consume();
        return { type: 'group', items };
      }
      if (ch === '\\') {
        return this._parseCommand();
      }
      if (ch) {
        return { type: 'text', text: this.consume() };
      }
      return { type: 'text', text: '' };
    }

    _parseCommand() {
      this.consume(); // skip \
      let cmd = '';
      while (this.pos < this.src.length && /[a-zA-Z]/.test(this.src[this.pos])) {
        cmd += this.consume();
      }
      if (!cmd) {
        // Single special char: \{ \} \\ \_ etc.
        const ch = this.consume();
        return { type: 'text', text: ch || '\\' };
      }

      const node = { type: 'command', cmd };

      // Fraction commands
      if (['frac', 'dfrac', 'tfrac'].includes(cmd)) {
        this.skipWs();
        node.num = this._parseSingle();
        this.skipWs();
        node.den = this._parseSingle();
        return node;
      }

      // Square root
      if (cmd === 'sqrt') {
        this.skipWs();
        if (this.peek() === '[') {
          this.consume();
          let deg = '';
          while (this.pos < this.src.length && this.peek() !== ']') {
            deg += this.consume();
          }
          if (this.peek() === ']') this.consume();
          node.degree = deg;
        }
        this.skipWs();
        node.content = this._parseSingle();
        return node;
      }

      // Nary operators
      if (cmd in NARY_OPS) {
        this.skipWs();
        if (this.peek() === '_') {
          this.consume();
          node.lower = this._parseSingle();
          this.skipWs();
        }
        if (this.peek() === '^') {
          this.consume();
          node.upper = this._parseSingle();
        }
        return node;
      }

      // Left delimiter
      if (cmd === 'left') {
        this.skipWs();
        node.delim = this._parseDelimiter();
        node.items = this._parseItemsStopAtRight();
        // Consume \right
        this.skipWs();
        if (this.pos + 5 < this.src.length && this.src.substring(this.pos, this.pos + 6) === '\\right') {
          this.pos += 6;
          this.skipWs();
          node.right_node = { delim: this._parseDelimiter() };
        }
        // Check for sub/sup on the whole delimited expression
        this.skipWs();
        if (this.peek() === '_') {
          this.consume();
          node.lower = this._parseSingle();
        }
        if (this.peek() === '^') {
          this.consume();
          node.upper = this._parseSingle();
        }
        return node;
      }

      // Accent commands
      if (cmd in ACCENT_CHARS) {
        this.skipWs();
        node.content = this._parseSingle();
        return node;
      }

      // Font commands
      if (cmd in FONT_STY) {
        this.skipWs();
        node.content = this._parseSingle();
        return node;
      }

      // overline / underline
      if (cmd === 'overline' || cmd === 'underline') {
        this.skipWs();
        node.content = this._parseSingle();
        return node;
      }

      // binom
      if (cmd === 'binom') {
        this.skipWs();
        node.num = this._parseSingle();
        this.skipWs();
        node.den = this._parseSingle();
        return node;
      }

      // stackrel
      if (cmd === 'stackrel') {
        this.skipWs();
        node.content = this._parseSingle();
        this.skipWs();
        node.upper = this._parseSingle();
        return node;
      }

      // over (infix fraction)
      if (cmd === 'over') {
        this.skipWs();
        node.content_right = this._parseSingle();
        return node;
      }

      // text, textrm, textit, textbf
      if (['text', 'textrm', 'textit', 'textbf'].includes(cmd)) {
        this.skipWs();
        node.content = this._parseSingle();
        return node;
      }

      // begin{env}
      if (cmd === 'begin') {
        this.skipWs();
        if (this.peek() === '{') {
          this.consume();
          let env = '';
          while (this.pos < this.src.length && this.peek() !== '}') {
            env += this.consume();
          }
          if (this.peek() === '}') this.consume();
          node.env = env;
          node.items = this._parseEnvItems(env);
          return node;
        }
        return node;
      }

      // end — should not be reached in normal flow
      if (cmd === 'end') {
        return node;
      }

      // right — should not be reached in normal flow
      if (cmd === 'right') {
        this.skipWs();
        if (this.pos < this.src.length) node.delim = this.consume();
        return node;
      }

      return node;
    }

    _parseDelimiter() {
      const ch = this.peek();
      if (ch === '\\') {
        this.consume();
        let cmd = '';
        while (this.pos < this.src.length && /[a-zA-Z]/.test(this.src[this.pos])) {
          cmd += this.consume();
        }
        return cmd || this.consume();
      }
      return this.consume();
    }

    _parseItemsStopAtRight() {
      const items = [];
      let iter = 0;
      while (this.pos < this.src.length && iter++ < this._maxIter) {
        const ch = this.peek();
        if (ch === '\\') {
          // Check for \right
          if (this.pos + 5 < this.src.length && this.src.substring(this.pos, this.pos + 6) === '\\right') {
            break;
          }
          items.push(this._parseCommand());
        } else if (ch === '}') {
          break;
        } else if (/\s/.test(ch)) {
          this.consume();
        } else {
          items.push({ type: 'text', text: this.consume() });
        }
      }
      return items;
    }

    _parseEnvItems(env) {
      const items = [];
      let iter = 0;
      const endMarker = `\\end{${env}}`;
      while (this.pos < this.src.length && iter++ < this._maxIter) {
        // Check for \end{env}
        if (this.src.substring(this.pos, this.pos + endMarker.length) === endMarker) {
          this.pos += endMarker.length;
          break;
        }

        const ch = this.peek();
        if (ch === '\\') {
          items.push(this._parseCommand());
        } else if (ch === '&') {
          this.consume();
          items.push({ type: 'amp' });
        } else if (ch === '\n' || (ch === '\\' && this.src[this.pos + 1] === '\\')) {
          if (ch === '\\') { this.consume(); this.consume(); }
          else this.consume();
          items.push({ type: 'newline' });
        } else if (ch === '{') {
          this.consume();
          const inner = this._parseItems('}');
          if (this.peek() === '}') this.consume();
          items.push(...inner);
        } else if (/\s/.test(ch)) {
          this.consume();
        } else {
          items.push({ type: 'text', text: this.consume() });
        }
      }
      return items;
    }
  }

  // ─── AST → OMML conversion ────────────────────────────────────────────────────

  function latexToOmml(latex) {
    const parser = new LaTeXParser(latex.trim());
    const tree = parser.parse();
    return ommlItems(tree.items || []);
  }

  function ommlNode(node) {
    if (!node || !node.type) return '';
    switch (node.type) {
      case 'text': return ommlText(node.text);
      case 'group': return ommlItems(node.items || []);
      case 'command': return ommlCommand(node);
      case 'sub': case 'sup': return ''; // handled by ommlItems
      default: return '';
    }
  }

  function ommlText(text) {
    if (!text) return '';
    return `<m:r><m:t>${esc(text)}</m:t></m:r>`;
  }

  function ommlItems(items) {
    const result = [];
    let i = 0;
    while (i < items.length) {
      const item = items[i];
      const itemType = item.type || '';

      // Subscript
      if (itemType === 'sub') {
        const base = popLast(result);
        const subContent = ommlNode(item.content || {});
        // Check if next is sup → combine as sSubSup
        if (i + 1 < items.length && items[i + 1].type === 'sup') {
          i++;
          const supContent = ommlNode(items[i].content || {});
          result.push(`<m:sSubSup><m:e>${base}</m:e><m:sub>${subContent}</m:sub><m:sup>${supContent}</m:sup></m:sSubSup>`);
        } else {
          result.push(`<m:sSub><m:e>${base}</m:e><m:sub>${subContent}</m:sub></m:sSub>`);
        }
        i++;
        continue;
      }

      // Superscript
      if (itemType === 'sup') {
        const base = popLast(result);
        const supContent = ommlNode(item.content || {});
        result.push(`<m:sSup><m:e>${base}</m:e><m:sup>${supContent}</m:sup></m:sSup>`);
        i++;
        continue;
      }

      // Nary operators: collect body items
      if (itemType === 'command' && NARY_OPS[item.cmd]) {
        let naryXml = ommlCommand(item);
        // Collect body: items following the nary operator until a boundary
        const bodyItems = [];
        i++;
        while (i < items.length) {
          const nextItem = items[i];
          const nextType = nextItem.type || '';
          if (nextType === 'command') {
            const nextCmd = nextItem.cmd || '';
            if (NARY_OPS[nextCmd]) break;
            if (['geq','ge','leq','le','eq','neq','ne','approx','equiv','sim','propto','left','right'].includes(nextCmd)) break;
          }
          if (nextType === 'text' && ['=','≥','≤','≠','>','<'].includes((nextItem.text || '').trim())) break;
          bodyItems.push(nextItem);
          i++;
        }
        const bodyXml = bodyItems.length > 0 ? ommlItems(bodyItems) : '';
        naryXml = naryXml.replace('<m:e></m:e>', `<m:e>${bodyXml}</m:e>`, 1);
        result.push(naryXml);
        continue;
      }

      result.push(ommlNode(item));
      i++;
    }
    return result.join('');
  }

  function popLast(result) {
    if (result.length === 0) return '<m:r><m:t></m:t></m:r>';
    return result.pop();
  }

  function ommlCommand(node) {
    const cmd = node.cmd || '';

    // Fractions
    if (['frac', 'dfrac', 'tfrac'].includes(cmd)) {
      const num = ommlNode(node.num || {});
      const den = ommlNode(node.den || {});
      let fPr = '';
      if (cmd === 'tfrac') fPr = '<m:fPr><m:type m:val="lin"/></m:fPr>';
      return `<m:f>${fPr}<m:num>${num}</m:num><m:den>${den}</m:den></m:f>`;
    }

    // Square root
    if (cmd === 'sqrt') {
      const content = ommlNode(node.content || {});
      const degree = node.degree;
      if (degree) {
        const degXml = ommlItems(new LaTeXParser(degree).parse().items || []);
        return `<m:rad><m:radPr><m:degHide m:val="0"/></m:radPr><m:deg>${degXml}</m:deg><m:e>${content}</m:e></m:rad>`;
      }
      return `<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e>${content}</m:e></m:rad>`;
    }

    // Nary operators
    if (NARY_OPS[cmd]) {
      const opChar = NARY_OPS[cmd];
      const hasLower = node.lower != null;
      const hasUpper = node.upper != null;
      const lowerXml = hasLower ? ommlNode(node.lower) : '';
      const upperXml = hasUpper ? ommlNode(node.upper) : '';
      const naryPr = `<m:naryPr><m:chr m:val="${esc(opChar)}"/><m:limLoc m:val="subSup"/><m:subHide m:val="${hasLower ? '0' : '1'}"/><m:supHide m:val="${hasUpper ? '0' : '1'}"/></m:naryPr>`;
      return `<m:nary>${naryPr}<m:sub>${lowerXml}</m:sub><m:sup>${upperXml}</m:sup><m:e></m:e></m:nary>`;
    }

    // Left delimiter
    if (cmd === 'left') {
      const delim = node.delim || '';
      const begChr = DELIM_MAP[delim] || delim;
      const inner = ommlItems(node.items || []);
      let endChr = '';
      if (node.right_node) {
        const endDelim = node.right_node.delim || '';
        endChr = DELIM_MAP[endDelim] || endDelim;
      }
      const dXml = `<m:d><m:dPr><m:begChr m:val="${esc(begChr)}"/><m:endChr m:val="${esc(endChr)}"/></m:dPr><m:e>${inner}</m:e></m:d>`;

      const lower = node.lower;
      const upper = node.upper;
      if (lower && upper) {
        return `<m:sSubSup><m:e>${dXml}</m:e><m:sub>${ommlNode(lower)}</m:sub><m:sup>${ommlNode(upper)}</m:sup></m:sSubSup>`;
      }
      if (upper) {
        return `<m:sSup><m:e>${dXml}</m:e><m:sup>${ommlNode(upper)}</m:sup></m:sSup>`;
      }
      if (lower) {
        return `<m:sSub><m:e>${dXml}</m:e><m:sub>${ommlNode(lower)}</m:sub></m:sSub>`;
      }
      return dXml;
    }

    // Accent
    if (ACCENT_CHARS[cmd]) {
      const content = ommlNode(node.content || {});
      const chr = ACCENT_CHARS[cmd];
      return `<m:acc><m:accPr><m:chr m:val="${chr}"/></m:accPr><m:e>${content}</m:e></m:acc>`;
    }

    // Font styling
    if (FONT_STY[cmd]) {
      const content = ommlNode(node.content || {});
      const sty = FONT_STY[cmd];
      // Apply styling to the first m:r in content
      return content.replace('<m:r>', `<m:r><m:rPr><m:sty m:val="${sty}"/></m:rPr>`, 1);
    }

    // Greek letters
    if (GREEK[cmd]) {
      return `<m:r><m:t>${GREEK[cmd]}</m:t></m:r>`;
    }

    // Symbols
    if (SYMBOLS[cmd]) {
      return `<m:r><m:t>${SYMBOLS[cmd]}</m:t></m:r>`;
    }

    // binom
    if (cmd === 'binom') {
      const num = ommlNode(node.num || {});
      const den = ommlNode(node.den || {});
      return `<m:d><m:dPr><m:begChr m:val="("/><m:endChr m:val=")"/></m:dPr><m:e><m:f><m:num>${num}</m:num><m:den>${den}</m:den></m:f></m:e></m:d>`;
    }

    // stackrel
    if (cmd === 'stackrel') {
      const content = ommlNode(node.content || {});
      const upper = ommlNode(node.upper || {});
      return `<m:sSup><m:e>${content}</m:e><m:sup>${upper}</m:sup></m:sSup>`;
    }

    // over (infix fraction — limited support)
    if (cmd === 'over') {
      const den = ommlNode(node.content_right || {});
      return `<m:f><m:num></m:num><m:den>${den}</m:den></m:f>`;
    }

    // text, textrm, textit, textbf
    if (['text', 'textrm', 'textit', 'textbf'].includes(cmd)) {
      const content = node.content || {};
      if (content.type === 'group') {
        const text = (content.items || []).map(it => it.text || '').join('');
        return `<m:r><m:t>${esc(text)}</m:t></m:r>`;
      }
      return ommlNode(content);
    }

    // begin{env} — matrix-like environments
    if (cmd === 'begin' && node.env) {
      return ommlEnv(node);
    }

    // overline / underline
    if (cmd === 'overline') {
      const content = ommlNode(node.content || {});
      return `<m:acc><m:accPr><m:chr m:val="0304"/></m:accPr><m:e>${content}</m:e></m:acc>`;
    }
    if (cmd === 'underline') {
      const content = ommlNode(node.content || {});
      return `<m:acc><m:accPr><m:chr m:val="0332"/></m:accPr><m:e>${content}</m:e></m:acc>`;
    }

    // Unknown command: output as text
    return `<m:r><m:t>${esc('\\' + cmd)}</m:t></m:r>`;
  }

  function ommlEnv(node) {
    const env = node.env || '';
    const items = node.items || [];

    // Split items into rows by newline, and cells by amp
    const rows = [];
    let currentRow = [];
    let currentCell = [];
    for (const item of items) {
      if (item.type === 'newline') {
        currentRow.push(currentCell);
        rows.push(currentRow);
        currentRow = [];
        currentCell = [];
      } else if (item.type === 'amp') {
        currentRow.push(currentCell);
        currentCell = [];
      } else {
        currentCell.push(item);
      }
    }
    if (currentCell.length > 0) currentRow.push(currentCell);
    if (currentRow.length > 0) rows.push(currentRow);

    // Build matrix XML
    const mc = rows.map(row => {
      const cells = row.map(cell => `<m:e>${ommlItems(cell)}</m:e>`).join('');
      return `<m:mr>${cells}</m:mr>`;
    }).join('');

    // Wrap with delimiters based on environment
    if (env === 'pmatrix') {
      return `<m:d><m:dPr><m:begChr m:val="("/><m:endChr m:val=")"/></m:dPr><m:e><m:m>${mc}</m:m></m:e></m:d>`;
    }
    if (env === 'bmatrix') {
      return `<m:d><m:dPr><m:begChr m:val="["/><m:endChr m:val="]"/></m:dPr><m:e><m:m>${mc}</m:m></m:e></m:d>`;
    }
    if (env === 'vmatrix') {
      return `<m:d><m:dPr><m:begChr m:val="|"/><m:endChr m:val="|"/></m:dPr><m:e><m:m>${mc}</m:m></m:e></m:d>`;
    }
    if (env === 'cases') {
      return `<m:d><m:dPr><m:begChr m:val="{"/><m:endChr m:val=" "/></m:dPr><m:e><m:m>${mc}</m:m></m:e></m:d>`;
    }
    // Plain matrix / aligned / split / array: no delimiters
    return `<m:d><m:dPr><m:begChr m:val=""/><m:endChr m:val=""/></m:dPr><m:e><m:m>${mc}</m:m></m:e></m:d>`;
  }

  function ommlWrap(inner, display) {
    const ns = `xmlns:m="${NS_M}"`;
    if (display) {
      return `<m:oMathPara ${ns}><m:oMath>${inner}</m:oMath></m:oMathPara>`;
    }
    return `<m:oMath ${ns}>${inner}</m:oMath>`;
  }

  // ═══════════════════════════════════════════════════════════════════════════════
  // Markdown-to-chunks parser
  // ═══════════════════════════════════════════════════════════════════════════════

  function extractMath(text) {
    const blocks = [];
    // Display math: $$...$$ or \[...\]
    let result = text;
    result = result.replace(/\$\$([\s\S]+?)\$\$/g, (_, latex) => {
      const idx = blocks.length;
      blocks.push({ type: 'math', latex: latex.trim(), display: true });
      return `\x00MATH${idx}\x00`;
    });
    result = result.replace(/\\\[([\s\S]+?)\\\]/g, (_, latex) => {
      const idx = blocks.length;
      blocks.push({ type: 'math', latex: latex.trim(), display: true });
      return `\x00MATH${idx}\x00`;
    });
    // Inline math: $...$
    result = result.replace(/\$([^\$\n]+?)\$/g, (_, latex) => {
      const idx = blocks.length;
      blocks.push({ type: 'math', latex: latex.trim(), display: false });
      return `\x00MATH${idx}\x00`;
    });
    return { text: result, blocks };
  }

  function parseMarkdown(text) {
    // First extract math expressions
    const { text: processedText, blocks: mathBlocks } = extractMath(text);

    const lines = processedText.split('\n');
    const chunks = [];
    let i = 0;

    while (i < lines.length) {
      const line = lines[i];

      // Empty line
      if (!line.trim()) { i++; continue; }

      // Code block start (```...)
      if (/^```/.test(line.trim())) {
        const lang = line.trim().substring(3).trim();
        const codeLines = [];
        i++;
        while (i < lines.length && !lines[i].trim().startsWith('```')) {
          codeLines.push(lines[i]);
          i++;
        }
        i++; // skip closing ```
        chunks.push({ type: 'code', language: lang, code: codeLines.join('\n') });
        continue;
      }

      // Table (|...|...|)
      if (/^\s*\|/.test(line) && /\|\s*$/.test(line.trim())) {
        const tableLines = [];
        while (i < lines.length && /^\s*\|/.test(lines[i]) && /\|\s*$/.test(lines[i].trim())) {
          tableLines.push(lines[i]);
          i++;
        }
        const table = parseMdTable(tableLines);
        if (table) chunks.push({ type: 'table', ...table });
        continue;
      }

      // Horizontal rule
      if (/^(-{3,}|_{3,}|\*{3,})\s*$/.test(line.trim())) {
        chunks.push({ type: 'hr' });
        i++;
        continue;
      }

      // Blockquote
      if (line.trim().startsWith('> ')) {
        const qLines = [];
        while (i < lines.length && lines[i].trim().startsWith('> ')) {
          qLines.push(lines[i].trim().substring(2));
          i++;
        }
        chunks.push({ type: 'blockquote', text: qLines.join('\n') });
        continue;
      }

      // Task list
      if (/^\s*[-*+]\s+\[[ xX]\]/.test(line)) {
        const items = [];
        while (i < lines.length && /^\s*[-*+]\s+\[[ xX]\]/.test(lines[i])) {
          const m = lines[i].match(/^\s*[-*+]\s+\[([ xX])\]\s*(.*)/);
          if (m) items.push({ checked: m[1].toLowerCase() === 'x', text: m[2] });
          i++;
        }
        chunks.push({ type: 'task_list', items });
        continue;
      }

      // Heading (# ...)
      const hMatch = line.match(/^(#{1,6})\s+(.*)/);
      if (hMatch) {
        chunks.push({ type: 'heading', level: hMatch[1].length, text: hMatch[2].trim() });
        i++;
        continue;
      }

      // Image ![alt](src)
      const imgMatch = line.trim().match(/^!\[([^\]]*)\]\(([^)]+)\)$/);
      if (imgMatch) {
        chunks.push({ type: 'image', alt: imgMatch[1], src: imgMatch[2] });
        i++;
        continue;
      }

      // Text paragraph (accumulate until block boundary)
      const textLines = [];
      while (i < lines.length) {
        const s = lines[i].trim();
        if (!s || /^```/.test(s) || /^\s*\|/.test(s) || /^(-{3,}|_{3,}|\*{3,})\s*$/.test(s) ||
            s.startsWith('> ') || /^\s*[-*+]\s+\[[ xX]\]/.test(s) ||
            /^(#{1,6})\s+/.test(s) || /^!\[([^\]]*)\]\(([^)]+)\)$/.test(s)) {
          break;
        }
        textLines.push(lines[i]);
        i++;
      }
      // Skip trailing empty lines at the end of text block
      while (i < lines.length && !lines[i].trim()) i++;
      const ct = textLines.join('\n').trim();
      if (ct) chunks.push({ type: 'text', text: ct });
    }

    // Replace math placeholders with actual math chunks
    const finalChunks = [];
    for (const chunk of chunks) {
      if (chunk.type === 'text') {
        // Split text by math placeholders
        const parts = chunk.text.split(/\x00MATH(\d+)\x00/);
        for (let j = 0; j < parts.length; j++) {
          if (j % 2 === 0) {
            // Text part
            if (parts[j].trim()) {
              finalChunks.push({ type: 'text', text: parts[j] });
            }
          } else {
            // Math block index
            const idx = parseInt(parts[j]);
            if (mathBlocks[idx]) {
              finalChunks.push(mathBlocks[idx]);
            }
          }
        }
      } else if (chunk.type === 'heading') {
        // Also process math in headings
        const parts = chunk.text.split(/\x00MATH(\d+)\x00/);
        const textParts = [];
        for (let j = 0; j < parts.length; j++) {
          if (j % 2 === 0) {
            textParts.push(parts[j]);
          } else {
            const idx = parseInt(parts[j]);
            if (mathBlocks[idx]) {
              textParts.push('$' + mathBlocks[idx].latex + '$');
            }
          }
        }
        finalChunks.push({ type: 'heading', level: chunk.level, text: textParts.join('') });
      } else {
        finalChunks.push(chunk);
      }
    }

    // Also add standalone math blocks that weren't inside text
    // (display math on its own line becomes a separate chunk)
    for (const block of mathBlocks) {
      if (block.display && !finalChunks.includes(block)) {
        // Check if this block was already included
        const isAlreadyIncluded = finalChunks.some(c => c === block);
        if (!isAlreadyIncluded) {
          finalChunks.push(block);
        }
      }
    }

    return finalChunks;
  }

  function parseMdTable(lines) {
    const rows = [];
    let headers = null;
    for (const line of lines) {
      const cells = line.trim().replace(/^\||\|$/g, '').split('|').map(c => c.trim());
      if (/^[\-:]+$/.test(cells[0])) continue; // separator row
      if (!headers) headers = cells;
      else rows.push(cells);
    }
    if (!headers) return null;
    return { headers, rows };
  }

  // ═══════════════════════════════════════════════════════════════════════════════
  // Document XML builder
  // ═══════════════════════════════════════════════════════════════════════════════

  function buildDocumentXml(chunks, title, options) {
    const bodyParts = [];
    const { toc, autoCaptions } = options || {};
    let figNum = 0, tblNum = 0;

    // Title
    if (title) {
      bodyParts.push(el('p', null,
        paraProps('center', 0, 360, null) +
        textRun(title, 'Arial', 22, true, false) +
        textRun(title, '黑体', 22, true, false)
      ));
    }

    // TOC placeholder
    if (toc) {
      bodyParts.push(el('p', null, paraProps('left', 0, 200) + el('r', null,
        el('fldChar', { 'w:fldCharType': 'begin' }) +
        el('instrText', { 'xml:space': 'preserve' }, ' TOC \\o "1-3" \\h \\z ') +
        el('fldChar', { 'w:fldCharType': 'separate' }) +
        el('r', null, el('t', null, '[ Right-click → Update Field in Word to generate TOC ]')) +
        el('fldChar', { 'w:fldCharType': 'end' })
      )));
    }

    for (const chunk of chunks) {
      const type = chunk.type || 'text';
      switch (type) {
        case 'heading': {
          const lvl = chunk.level || 1;
          const runs = runsFromText(chunk.text, 'Arial', [22, 18, 16, 14, 13, 12][Math.min(lvl, 6) - 1] || 12);
          const styleName = `Heading${lvl}`;
          bodyParts.push(el('p', null,
            el('pPr', null, el('pStyle', { 'w:val': styleName })) +
            runs.join('')
          ));
          break;
        }
        case 'text': {
          const runs = runsFromText(chunk.text, 'Times New Roman', 12);
          bodyParts.push(el('p', null,
            paraProps(null, 0, 0, 360) +
            runs.join('')
          ));
          break;
        }
        case 'table': {
          if (autoCaptions) tblNum++;
          if (autoCaptions) {
            bodyParts.push(el('p', null,
              paraProps('center', 240, 120, 360) +
              textRun(`表${tblNum} ${chunk.caption || ''}`, '宋体', 10, true, false)
            ));
          }
          const headers = chunk.headers || [];
          const rows = chunk.rows || [];
          const cols = headers.length || 1;
          const tblParts = [];
          tblParts.push(el('tblPr', null,
            el('tblStyle', { 'w:val': 'TableGrid' }) +
            el('tblW', { 'w:w': '5000', 'w:type': 'pct' }) +
            tableBorders()
          ));
          const hdrCells = headers.map(h =>
            el('tc', null,
              el('tcPr', null, cellShading('D9E2F3')) +
              el('p', null, paraProps('left', 0, 0, 276) + textRun(h, 'Times New Roman', 10, true, false))
            )
          );
          tblParts.push(el('tr', null, hdrCells.join('')));
          for (const row of rows) {
            const rowCells = [];
            for (let c = 0; c < cols; c++) {
              const cellText = c < row.length ? row[c] : '';
              rowCells.push(el('tc', null,
                el('p', null, paraProps('left', 0, 0, 276) + textRun(cellText, 'Times New Roman', 10, false, false))
              ));
            }
            tblParts.push(el('tr', null, rowCells.join('')));
          }
          bodyParts.push(el('tbl', null, tblParts.join('')));
          bodyParts.push(el('p', null, emptyRun()));
          break;
        }
        case 'code': {
          const codeLines = (chunk.code || '').split('\n');
          for (const cl of codeLines) {
            bodyParts.push(el('p', null,
              shadedParaProps('left', 0, 0, 'F5F5F5') +
              textRun(cl || ' ', 'Consolas', 9, false, false)
            ));
          }
          bodyParts.push(el('p', null, emptyRun()));
          break;
        }
        case 'task_list': {
          const items = chunk.items || [];
          for (const item of items) {
            const chk = item.checked ? '☒' : '☐';
            bodyParts.push(el('p', null,
              paraProps('left', 0, 0, 360) +
              textRun(`${chk}  ${item.text}`, 'Times New Roman', 12, false, false)
            ));
          }
          break;
        }
        case 'blockquote': {
          bodyParts.push(el('p', null,
            shadedParaProps('left', 120, 120, 'F0F0F0', 720) +
            runsFromText(chunk.text, 'Times New Roman', 10.5).join('')
          ));
          break;
        }
        case 'hr': {
          const pBdr = el('pBdr', null, el('bottom', { 'w:val': 'single', 'w:sz': '6', 'w:space': '1', 'w:color': '999999' }));
          bodyParts.push(el('p', null, el('pPr', null, el('spacing', { 'w:before': '120', 'w:after': '120' }) + pBdr)));
          break;
        }
        case 'image': {
          if (autoCaptions) figNum++;
          if (autoCaptions) {
            bodyParts.push(el('p', null,
              paraProps('center', 120, 240, 360) +
              textRun(`图${figNum} ${chunk.alt || ''}`, '宋体', 10, false, false)
            ));
          }
          bodyParts.push(el('p', null,
            paraProps('center', 0, 0, 360) +
            textRun(`[Image: ${chunk.alt || chunk.src || ''}]`, 'Times New Roman', 10, false, false, '999999')
          ));
          bodyParts.push(el('p', null, emptyRun()));
          break;
        }
        case 'math': {
          try {
            const latex = chunk.latex || '';
            const display = chunk.display || false;
            const innerOmml = latexToOmml(latex);
            const wrappedXml = ommlWrap(innerOmml, display);
            if (display) {
              bodyParts.push(`<w:p>${paraProps('center', 60, 60, 360)}${wrappedXml}</w:p>`);
            } else {
              // Inline math: standalone paragraph for simplicity
              bodyParts.push(`<w:p>${wrappedXml}</w:p>`);
            }
          } catch (e) {
            // Fallback to styled text
            const latex = chunk.latex || '';
            bodyParts.push(el('p', null,
              paraProps(chunk.display ? 'center' : 'left', 60, 60, 360) +
              textRun(`$${latex}$`, 'Cambria Math', 11, false, true, '333333')
            ));
          }
          break;
        }
      }
    }

    return bodyParts.join('');
  }

  function buildStylesXml() {
    const styles = [];
    styles.push(el('style', { 'w:type': 'paragraph', 'w:styleId': 'Normal', 'w:default': '1' },
      el('name', { 'w:val': 'Normal' }) +
      el('rPr', null,
        el('rFonts', { 'w:ascii': 'Times New Roman', 'w:hAnsi': 'Times New Roman', 'w:eastAsia': '宋体' }) +
        el('sz', { 'w:val': '24' })) +
      el('pPr', null, el('spacing', { 'w:line': '360', 'w:lineRule': 'auto' }))
    ));

    const hSizes = [36, 32, 28, 24, 22, 20];
    for (let l = 1; l <= 6; l++) {
      styles.push(el('style', { 'w:type': 'paragraph', 'w:styleId': `Heading${l}` },
        el('name', { 'w:val': `heading ${l}` }) +
        el('basedOn', { 'w:val': 'Normal' }) +
        el('next', { 'w:val': 'Normal' }) +
        el('rPr', null,
          el('rFonts', { 'w:ascii': 'Arial', 'w:hAnsi': 'Arial', 'w:eastAsia': '黑体' }) +
          el('b') +
          el('sz', { 'w:val': String(hSizes[l - 1]) })) +
        el('pPr', null,
          el('outlineLvl', { 'w:val': String(l - 1) }) +
          el('spacing', { 'w:before': String((7 - l) * 60), 'w:after': String((7 - l) * 30) }))
      ));
    }

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
${styles.join('\n')}
</w:styles>`;
  }

  // ─── Minimal ZIP creator (STORE method only) ──────────────────────────────────

  function createZip(files) {
    const enc = new TextEncoder();
    const entries = [];
    let offset = 0;

    for (const f of files) {
      const data = f.data instanceof Uint8Array ? f.data : enc.encode(f.data);
      const nameBytes = enc.encode(f.name);
      const lh = new Uint8Array(30 + nameBytes.length);
      const lv = new DataView(lh.buffer);
      lv.setUint32(0, 0x04034b50, true);
      lv.setUint16(8, 0, true);
      lv.setUint32(14, crc32(data), true);
      lv.setUint32(18, data.length, true);
      lv.setUint32(22, data.length, true);
      lv.setUint16(26, nameBytes.length, true);
      lh.set(nameBytes, 30);
      entries.push({ lh, data, nameBytes, offset });
      offset += lh.length + data.length;
    }

    let cdsize = 0;
    const cdParts = [];
    for (const e of entries) {
      const cd = new Uint8Array(46 + e.nameBytes.length);
      const cv = new DataView(cd.buffer);
      cv.setUint32(0, 0x02014b50, true);
      cv.setUint16(8, 0, true);
      cv.setUint16(10, 0, true);
      cv.setUint16(12, 0, true);
      cv.setUint16(14, 0, true);
      cv.setUint32(16, crc32(e.data), true);
      cv.setUint32(20, e.data.length, true);
      cv.setUint32(24, e.data.length, true);
      cv.setUint16(28, e.nameBytes.length, true);
      cv.setUint32(42, e.offset, true);
      cd.set(e.nameBytes, 46);
      cdParts.push(cd);
      cdsize += cd.length;
    }

    const eocd = new Uint8Array(22);
    const ev = new DataView(eocd.buffer);
    ev.setUint32(0, 0x06054b50, true);
    ev.setUint16(8, entries.length, true);
    ev.setUint16(10, entries.length, true);
    ev.setUint32(12, cdsize, true);
    ev.setUint32(16, offset, true);

    const total = offset + cdsize + 22;
    const result = new Uint8Array(total);
    let pos = 0;
    for (const e of entries) {
      result.set(e.lh, pos); pos += e.lh.length;
      result.set(e.data, pos); pos += e.data.length;
    }
    for (const cd of cdParts) {
      result.set(cd, pos); pos += cd.length;
    }
    result.set(eocd, pos);
    return result;
  }

  function crc32(data) {
    let crc = 0xFFFFFFFF;
    for (let i = 0; i < data.length; i++) {
      crc ^= data[i];
      for (let j = 0; j < 8; j++) {
        crc = (crc >>> 1) ^ (crc & 1 ? 0xEDB88320 : 0);
      }
    }
    return (crc ^ 0xFFFFFFFF) >>> 0;
  }

  // ─── Content types & rels XML ─────────────────────────────────────────────────

  function buildContentTypes() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`;
  }

  function buildRels() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;
  }

  function buildDocRels() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;
  }

  // ─── Public API ───────────────────────────────────────────────────────────────

  function buildDocx(content, options = {}) {
    const chunks = parseMarkdown(content);
    const title = options.title || '';
    const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <w:body>
${buildDocumentXml(chunks, title, options)}
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>
    </w:sectPr>
  </w:body>
</w:document>`;

    const files = [
      { name: '[Content_Types].xml', data: buildContentTypes() },
      { name: '_rels/.rels', data: buildRels() },
      { name: 'word/document.xml', data: docXml },
      { name: 'word/styles.xml', data: buildStylesXml() },
      { name: 'word/_rels/document.xml.rels', data: buildDocRels() },
    ];

    return createZip(files);
  }

  function getStats(content) {
    const chunks = parseMarkdown(content);
    const stats = {};
    for (const c of chunks) {
      const t = c.type || 'text';
      stats[t] = (stats[t] || 0) + 1;
    }
    return { chunks, stats };
  }

  return { buildDocx, getStats, parseMarkdown, latexToOmml };
})();

// Export for use in service worker / popup contexts
if (typeof module !== 'undefined' && module.exports) {
  module.exports = DocxBuilder;
}
