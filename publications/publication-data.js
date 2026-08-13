(function (global) {
  'use strict';

  const CATEGORY_ALIASES = {
    hci: 'hci',
    'hci / ai': 'hci',
    'hci/ai': 'hci',
    ai: 'hci',
    '人机交互': 'hci',
    '人工智能': 'hci',
    emergency: 'emergency',
    '应急': 'emergency',
    '应急管理': 'emergency',
    av: 'av',
    '自动驾驶': 'av',
    '人车交互': 'av',
    biblio: 'biblio',
    bibliometrics: 'biblio',
    '文献计量': 'biblio'
  };

  function cleanText(value) {
    return value == null ? '' : String(value).trim();
  }

  function splitTags(value) {
    return cleanText(value)
      .split(/[;；,，]/)
      .map(tag => tag.trim())
      .filter(Boolean);
  }

  function normalizeDoi(value) {
    const doi = cleanText(value).replace(/\s+/g, '');
    if (!doi) return '';

    if (/^https?:\/\//i.test(doi)) return doi;

    const identifier = doi.replace(/^doi:/i, '').replace(/^https?:\/\/(?:dx\.)?doi\.org\//i, '');
    return identifier ? `https://doi.org/${identifier}` : '';
  }

  function normalizeCategory(value, row) {
    const explicit = cleanText(value).toLowerCase();
    if (CATEGORY_ALIASES[explicit]) return CATEGORY_ALIASES[explicit];

    const text = [row['论文标题'], row['出版物'], row['标签']]
      .map(cleanText)
      .join(' ')
      .toLowerCase();

    if (/自动驾驶|自主车辆|人车交互|感知安全|automated driv|automated vehicle|autonomous vehicle/.test(text)) {
      return 'av';
    }
    if (/文献计量|bibliometric|scientometric/.test(text)) {
      return 'biblio';
    }
    if (/应急|灾害|预警|风险|消防|火灾|emergency|disaster|warning|risk perception|fire monitoring|natural disaster/.test(text)) {
      return 'emergency';
    }
    return 'hci';
  }

  function getFirstValue(row, keys) {
    for (const key of keys) {
      if (row[key] != null && cleanText(row[key])) return row[key];
    }
    return '';
  }

  function mapRow(row, sourceIndex) {
    const title = cleanText(getFirstValue(row, ['论文标题', '标题', 'title', 'Title']));
    if (!title) return null;

    const categoryValue = getFirstValue(row, ['分类', '类别', 'category', 'cat']);
    const yearValue = getFirstValue(row, ['年份', '年', 'year', 'Year']);
    const parsedYear = Number.parseInt(cleanText(yearValue), 10);

    return {
      title,
      venue: cleanText(getFirstValue(row, ['出版物', '期刊/会议', '期刊', 'venue', 'Venue'])),
      year: Number.isFinite(parsedYear) ? parsedYear : cleanText(yearValue),
      authors: cleanText(getFirstValue(row, ['全部作者', '作者', 'authors', 'Authors'])),
      index: cleanText(getFirstValue(row, ['索引', '收录', 'index', 'Index'])),
      tags: splitTags(getFirstValue(row, ['标签', '关键词', 'tags', 'Tags'])),
      doi: normalizeDoi(getFirstValue(row, ['doi地址', 'DOI地址', 'DOI', 'doi'])),
      cat: normalizeCategory(categoryValue, row),
      _sourceIndex: sourceIndex
    };
  }

  async function load(url) {
    if (!global.XLSX) {
      throw new Error('Excel 解析器未加载');
    }

    let response;
    try {
      response = await fetch(url, { cache: 'no-cache' });
    } catch (error) {
      if (global.location && global.location.protocol === 'file:') {
        throw new Error('浏览器不允许网页直接读取本地 Excel，请通过本地网站服务打开主页');
      }
      throw error;
    }

    if (!response.ok) {
      throw new Error(`读取论文清单失败（HTTP ${response.status}）`);
    }

    const bytes = await response.arrayBuffer();
    const workbook = global.XLSX.read(bytes, { type: 'array' });
    const firstSheetName = workbook.SheetNames[0];
    if (!firstSheetName) throw new Error('论文清单中没有工作表');

    const rows = global.XLSX.utils.sheet_to_json(workbook.Sheets[firstSheetName], {
      defval: '',
      raw: false
    });

    return rows
      .map(mapRow)
      .filter(Boolean)
      .sort((a, b) => {
        const yearDifference = (Number(b.year) || 0) - (Number(a.year) || 0);
        return yearDifference || a._sourceIndex - b._sourceIndex;
      })
      .map(({ _sourceIndex, ...publication }) => publication);
  }

  function escapeHtml(value) {
    return cleanText(value).replace(/[&<>'"]/g, char => ({
      '&': '&amp;',
      '<': '&lt;',
      '>': '&gt;',
      "'": '&#39;',
      '"': '&quot;'
    })[char]);
  }

  global.PublicationData = { load, escapeHtml };
})(window);
