const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, PageOrientation, LevelFormat,
  TabStopType, TabStopPosition, HeadingLevel, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, PageBreak
} = require('docx');

// ===== Style Constants =====
const FONT = 'Arial';
const COLOR_PRIMARY = '1A6B5C';
const COLOR_HEADER_BG = 'D5E8F0';
const COLOR_ALT_BG = 'F0F4F8';
const COLOR_CODE_BG = 'F5F5F5';
const COLOR_MUTED = '666666';
const BORDER = { style: BorderStyle.SINGLE, size: 1, color: 'BFBFBF' };
const BORDERS = { top: BORDER, bottom: BORDER, left: BORDER, right: BORDER };
const NO_BORDER = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const NO_BORDERS = { top: NO_BORDER, bottom: NO_BORDER, left: NO_BORDER, right: NO_BORDER };
const CELL_MARGINS = { top: 60, bottom: 60, left: 100, right: 100 };

// US Letter content width with 1" margins
const CONTENT_WIDTH = 9360;

// ===== Helper Functions =====
function txt(text, opts = {}) {
  return new TextRun({ text, font: FONT, ...opts });
}

function bold(text, opts = {}) {
  return txt(text, { bold: true, ...opts });
}

function para(children, opts = {}) {
  return new Paragraph({ children, ...opts });
}

function textPara(text, opts = {}) {
  return para([txt(text)], opts);
}

function headerCell(text, width, opts = {}) {
  return new TableCell({
    borders: BORDERS,
    width: { size: width, type: WidthType.DXA },
    shading: { fill: COLOR_HEADER_BG, type: ShadingType.CLEAR },
    margins: CELL_MARGINS,
    verticalAlign: VerticalAlign.CENTER,
    children: [para([bold(text)], { alignment: AlignmentType.CENTER })]
  });
}

function dataCell(text, width, opts = {}) {
  const isAlt = opts.alt || false;
  return new TableCell({
    borders: BORDERS,
    width: { size: width, type: WidthType.DXA },
    shading: isAlt ? { fill: COLOR_ALT_BG, type: ShadingType.CLEAR } : undefined,
    margins: CELL_MARGINS,
    verticalAlign: VerticalAlign.CENTER,
    children: [para([txt(text)], { alignment: opts.align || AlignmentType.LEFT })]
  });
}

function makeTable(headers, rows, colWidths) {
  const totalWidth = colWidths.reduce((a, b) => a + b, 0);
  const headerRow = new TableRow({
    tableHeader: true,
    children: headers.map((h, i) => headerCell(h, colWidths[i]))
  });
  const dataRows = rows.map((row, ri) =>
    new TableRow({
      children: row.map((cell, ci) => dataCell(cell, colWidths[ci], { alt: ri % 2 === 1 }))
    })
  );
  return new Table({
    width: { size: totalWidth, type: WidthType.DXA },
    columnWidths: colWidths,
    rows: [headerRow, ...dataRows]
  });
}

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 360, after: 200 },
    children: [new TextRun({ text, font: FONT, size: 36, bold: true, color: COLOR_PRIMARY })]
  });
}

function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 280, after: 160 },
    children: [new TextRun({ text, font: FONT, size: 28, bold: true, color: COLOR_PRIMARY })]
  });
}

function h3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    spacing: { before: 200, after: 120 },
    children: [new TextRun({ text, font: FONT, size: 24, bold: true, color: '333333' })]
  });
}

function h4(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_4,
    spacing: { before: 160, after: 100 },
    children: [new TextRun({ text, font: FONT, size: 22, bold: true, color: '444444' })]
  });
}

function blockquote(text) {
  return new Paragraph({
    spacing: { before: 100, after: 100 },
    indent: { left: 360 },
    border: { left: { style: BorderStyle.SINGLE, size: 6, color: COLOR_PRIMARY, space: 8 } },
    children: [new TextRun({ text, font: FONT, size: 22, italics: true, color: '555555' })]
  });
}

function numberedList(items) {
  return items.map((item, i) =>
    new Paragraph({
      numbering: { reference: 'numList', level: 0 },
      spacing: { before: 40, after: 40 },
      children: [txt(item)]
    })
  );
}

function bulletList(items) {
  return items.map(item =>
    new Paragraph({
      numbering: { reference: 'bullets', level: 0 },
      spacing: { before: 40, after: 40 },
      children: [txt(item)]
    })
  );
}

function codeBlock(lines) {
  return lines.map(line =>
    new Paragraph({
      spacing: { before: 0, after: 0 },
      shading: { fill: COLOR_CODE_BG, type: ShadingType.CLEAR },
      children: [new TextRun({ text: line || ' ', font: 'Courier New', size: 18 })]
    })
  );
}

function spacer() {
  return new Paragraph({ spacing: { before: 80, after: 80 }, children: [txt('')] });
}

// ===== Build Document =====
const doc = new Document({
  styles: {
    default: {
      document: { run: { font: FONT, size: 22 } }
    },
    paragraphStyles: [
      {
        id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 36, bold: true, font: FONT, color: COLOR_PRIMARY },
        paragraph: { spacing: { before: 360, after: 200 }, outlineLevel: 0 }
      },
      {
        id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 28, bold: true, font: FONT, color: COLOR_PRIMARY },
        paragraph: { spacing: { before: 280, after: 160 }, outlineLevel: 1 }
      },
      {
        id: 'Heading3', name: 'Heading 3', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 24, bold: true, font: FONT, color: '333333' },
        paragraph: { spacing: { before: 200, after: 120 }, outlineLevel: 2 }
      },
      {
        id: 'Heading4', name: 'Heading 4', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 22, bold: true, font: FONT, color: '444444' },
        paragraph: { spacing: { before: 160, after: 100 }, outlineLevel: 3 }
      }
    ]
  },
  numbering: {
    config: [
      {
        reference: 'bullets',
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: '\u2022',
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } }
        }]
      },
      {
        reference: 'numList',
        levels: [{
          level: 0, format: LevelFormat.DECIMAL, text: '%1.',
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } }
        }]
      }
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
      }
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          alignment: AlignmentType.RIGHT,
          children: [new TextRun({ text: 'PLA 丝绸质感 Filament 实验方案', font: FONT, size: 18, color: '999999' })]
        })]
      })
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({ text: '第 ', font: FONT, size: 18, color: '999999' }),
            new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 18, color: '999999' }),
            new TextRun({ text: ' 页', font: FONT, size: 18, color: '999999' })
          ]
        })]
      })
    },
    children: [
      // ===== TITLE =====
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 400, after: 200 },
        children: [new TextRun({ text: 'PLA 丝绸质感 Filament 实验方案', font: FONT, size: 44, bold: true, color: COLOR_PRIMARY })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 300 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: COLOR_PRIMARY, space: 4 } },
        children: [new TextRun({ text: '采购清单 \u00B7 实验矩阵 \u00B7 工艺参数 \u00B7 评价体系 \u00B7 操作流程', font: FONT, size: 20, color: COLOR_MUTED })]
      }),

      // ===== Intro Blockquote =====
      blockquote('目标：通过 PLA 基料 + PBAT 弹性体 + PEG 增塑剂 + 珠光粉的组合实验，找到丝绸光泽最佳的配方比例和工艺参数。'),
      blockquote('实验周期：约 2-3 周（含材料采购、混料、挤出、测试分析）'),

      spacer(),

      // ===== Section 1: Procurement =====
      h1('一、采购清单'),

      h2('1. 原材料'),
      makeTable(
        ['序号', '名称', '规格', '推荐品牌/供应商', '数量', '预估单价', '预估总价', '用途'],
        [
          ['1', 'PLA 颗粒', '4032D（注塑级）', 'NatureWorks / 中粮', '5 kg', '¥80/kg', '¥400', '基础树脂'],
          ['2', 'PBAT 颗粒', 'TH801（融指3-5）', '蓝山屯河 / 金发科技', '2 kg', '¥35/kg', '¥70', '弹性体改性，丝绸光泽'],
          ['3', 'PEG-8000', '化学纯，MW 8000', '国药集团 / 阿拉丁', '500 g', '¥45/瓶', '¥45', '增塑剂，改善流动性'],
          ['4', '珠光粉（银白）', '100目，银白系列', '默克 Iriodin 100', '100 g', '¥60/瓶', '¥60', '珠光光泽效果'],
          ['5', '珠光粉（金色）', '100目，金铜系列', '默克 Iriodin 300', '50 g', '¥70/瓶', '¥70', '金色丝绸变体'],
          ['6', 'POE-g-MA', '熔指 1-2g/10min', '陶氏 / 上海日之升', '500 g', '¥50/kg', '¥25', '相容剂'],
          ['7', '抗氧剂 1010', '化学纯', '巴斯夫 / 国药', '100 g', '¥35/瓶', '¥35', '防止热降解'],
          ['8', '色粉（可选）', '各色', '万华 / 道尔', '各 20 g', '¥15/色', '¥60', '染色丝绸变体']
        ],
        [600, 1200, 1400, 1600, 700, 800, 800, 1260]
      ),
      para([bold('原材料小计：约 ¥765')], { spacing: { before: 100, after: 200 } }),

      h2('2. 实验耗材'),
      makeTable(
        ['序号', '名称', '规格', '数量', '预估单价', '预估总价'],
        [
          ['1', '真空包装袋', '20×30cm 铝箔袋', '50 个', '¥0.5', '¥25'],
          ['2', '干燥剂', '硅胶 10g 包', '100 包', '¥0.1', '¥10'],
          ['3', '标签纸', '防水 5×3cm', '200 张', '¥0.05', '¥10'],
          ['4', '记号笔', '油性防水', '3 支', '¥5', '¥15'],
          ['5', '一次性手套', '丁腈手套', '1 盒', '¥25', '¥25'],
          ['6', '电子秤', '0.01g 精度', '1 台', '¥120', '¥120'],
          ['7', '游标卡尺', '0-150mm，0.02mm', '1 把', '¥50', '¥50'],
          ['8', '密封混合容器', '500ml 不锈钢盆', '5 个', '¥15', '¥75']
        ],
        [600, 2000, 2200, 1200, 1560, 1560]
      ),
      para([bold('耗材小计：约 ¥330')], { spacing: { before: 100, after: 200 } }),

      h2('3. 设备（如已有可跳过）'),
      makeTable(
        ['序号', '名称', '规格', '备注'],
        [
          ['1', '单/双螺杆挤出机', '螺杆直径 20-30mm，L/D ≥ 25', '双螺杆更佳，单螺杆也可'],
          ['2', '恒温干燥箱', '60°C，带鼓风', 'PLA 干燥除湿用'],
          ['3', '水槽冷却', '常温水循环', '牵引冷却'],
          ['4', '牵引机', '速度可调 10-100 mm/s', ''],
          ['5', '绕线机', '自动排线', ''],
          ['6', '3D 打印机', '0.4mm 喷嘴', '打印测试样件'],
          ['7', '光泽度计', '60° 角', '量化光泽度（可选，约¥300）'],
          ['8', '拉力试验机', '—', '测层间结合力（可选，约¥500）']
        ],
        [600, 2200, 2600, 3760]
      ),

      // ===== Section 2: Experiment Design =====
      h1('二、实验方案设计'),

      h2('实验矩阵（正交实验法）'),
      para([txt('采用 '), bold('L9（3^4）正交表'), txt('，4 因素 3 水平，共 9 组实验 + 3 组验证实验 = '), bold('12 组'), txt('。')]),

      h3('因素与水平'),
      makeTable(
        ['水平', 'A: PBAT (%)', 'B: PEG (%)', 'C: 珠光粉 (%)', 'D: 相容剂 (%)'],
        [
          ['1', '5', '1', '0', '0'],
          ['2', '10', '2', '1', '1'],
          ['3', '15', '3', '2', '2']
        ],
        [1200, 2040, 2040, 2040, 2040]
      ),

      h3('实验分组表'),
      makeTable(
        ['实验号', 'A: PBAT', 'B: PEG', 'C: 珠光粉', 'D: 相容剂', 'PLA 余量', '配方标识'],
        [
          ['E1', '5%', '1%', '0%', '0%', '94%', 'S05-1-0-0'],
          ['E2', '5%', '2%', '1%', '1%', '91%', 'S05-2-1-1'],
          ['E3', '5%', '3%', '2%', '2%', '88%', 'S05-3-2-2'],
          ['E4', '10%', '1%', '1%', '2%', '86%', 'S10-1-1-2'],
          ['E5', '10%', '2%', '2%', '0%', '86%', 'S10-2-2-0'],
          ['E6', '10%', '3%', '0%', '1%', '86%', 'S10-3-0-1'],
          ['E7', '15%', '1%', '2%', '1%', '81%', 'S15-1-2-1'],
          ['E8', '15%', '2%', '0%', '2%', '81%', 'S15-2-0-2'],
          ['E9', '15%', '3%', '1%', '0%', '81%', 'S15-3-1-0']
        ],
        [900, 1200, 1200, 1200, 1200, 1200, 2460]
      ),

      h3('验证实验（基于正交分析最优组合 + 对照组）'),
      makeTable(
        ['实验号', '配方', '目的'],
        [
          ['V1', '正交分析得出的最优组合', '验证最优配方重现性'],
          ['V2', '纯 PLA（无任何添加剂）', '空白对照'],
          ['V3', 'PLA + 10% PBAT（仅基础改性）', '验证珠光粉/PEG 单独贡献']
        ],
        [900, 3600, 4860]
      ),

      // ===== Section 3: Process Parameters =====
      h1('三、工艺参数（固定条件）'),
      para([txt('所有实验组使用统一的工艺参数（消除工艺变量干扰）：')]),
      makeTable(
        ['参数', '设定值', '说明'],
        [
          ['预干燥', 'PLA: 60°C × 4h；PBAT: 50°C × 3h', '除湿至含水率 <0.1%'],
          ['料桶温度', '195°C', '比 PLA 单独加工高 15°C，保证 PBAT 充分熔融'],
          ['螺杆速度', '26 RPM', '降速保证混炼均匀'],
          ['牵引速度', '32 mm/s', '与螺杆速度匹配，控制直径 1.75mm'],
          ['冷却水槽温度', '常温（25°C）', ''],
          ['抗氧剂 1010', '固定添加 0.2%', '所有组统一']
        ],
        [1800, 2800, 4760]
      ),

      // ===== Section 4: Evaluation =====
      h1('四、评价体系'),

      h2('1. 感官评价（目视+触感）'),
      makeTable(
        ['指标', '评分 1-5 分', '标准'],
        [
          ['光泽度', '1=哑光 5=镜面', '60°光泽度计辅助'],
          ['丝绸感', '1=无 5=强烈丝绸质感', '与市售 silk PLA 对比'],
          ['表面光滑度', '1=粗糙 5=极光滑', '手指触摸 + 放大镜观察'],
          ['颜色均匀性', '1=斑驳 5=均匀', '目视']
        ],
        [2000, 2800, 4560]
      ),

      h2('2. 力学性能（可选）'),
      makeTable(
        ['指标', '测试方法', '目标值'],
        [
          ['拉伸强度', 'ASTM D638', '≥ 30 MPa'],
          ['断裂伸长率', 'ASTM D638', '≥ 5%'],
          ['层间结合力', '打印 10×10×10mm 方块，手撕测试', '≥ 普通 PLA 80%']
        ],
        [2000, 4000, 3360]
      ),

      h2('3. 打印测试'),
      para([txt('打印标准化测试模型（温度塔 + 光泽样片），评价：')]),
      ...bulletList([
        '190°C / 200°C / 210°C 三档温度打印',
        '外墙光泽均匀性',
        '层附着力'
      ]),

      h2('4. 直径一致性'),
      makeTable(
        ['指标', '测量方法', '合格范围'],
        [
          ['直径', '游标卡尺，每隔 1m 测一次', '1.70-1.80mm'],
          ['椭圆度', '同截面测两个方向', '差值 <0.05mm']
        ],
        [2000, 4000, 3360]
      ),

      // ===== Section 5: Record Sheet =====
      h1('五、实验记录表'),

      h2('单次实验记录卡'),
      ...codeBlock([
        '===============================================',
        '  PLA 丝绸实验记录卡',
        '===============================================',
        '',
        '实验编号：______  日期：______  操作人：______',
        '',
        '■ 配方',
        '  PLA:        ______g  (______%)',
        '  PBAT:       ______g  (______%)',
        '  PEG-8000:   ______g  (______%)',
        '  珠光粉:      ______g  (______%)',
        '  POE-g-MA:   ______g  (______%)',
        '  抗氧剂1010:  ______g  (0.2%)',
        '',
        '■ 预干燥',
        '  PLA:  60°C × ______h  含水率: ______%',
        '  PBAT: 50°C × ______h  含水率: ______%',
        '',
        '■ 工艺参数',
        '  料桶温度:   ______°C',
        '  螺杆速度:   ______RPM',
        '  牵引速度:   ______mm/s',
        '  水槽温度:   ______°C',
        '',
        '■ 挤出观察',
        '  熔体状态:   □ 光滑 □ 粗糙 □ 有气泡 □ 有条纹',
        '  出料连续性: □ 连续 □ 断续 □ 偶尔断裂',
        '  备注:       _________________________________',
        '',
        '■ 直径测量（每米测一次，取 10 点）',
        '  平均值: ______mm  最大: ______mm  最小: ______mm',
        '  椭圆度: ______mm',
        '',
        '■ 感官评分',
        '  光泽度:     □1 □2 □3 □4 □5',
        '  丝绸感:     □1 □2 □3 □4 □5',
        '  表面光滑度: □1 □2 □3 □4 □5',
        '  颜色均匀性: □1 □2 □3 □4 □5',
        '',
        '■ 打印测试',
        '  打印温度: ______°C',
        '  外墙光泽: □均匀 □不均匀 □哑光线',
        '  层附着力: □良好 □一般 □差',
        '  备注:     _________________________________',
        '',
        '■ 总评',
        '  综合评分: ______/20',
        '  是否进入下一轮: □ 是 □ 否',
        '  改进建议: _________________________________',
        '',
        '==============================================='
      ]),

      h2('实验汇总表（9组正交 + 3组验证）'),
      makeTable(
        ['编号', 'PBAT%', 'PEG%', '珠光%', '相容剂%', '光泽', '丝绸感', '光滑', '均匀', '总分/20', '直径mm', '备注'],
        [
          ['E1', '5', '1', '0', '0', '', '', '', '', '', '', ''],
          ['E2', '5', '2', '1', '1', '', '', '', '', '', '', ''],
          ['E3', '5', '3', '2', '2', '', '', '', '', '', '', ''],
          ['E4', '10', '1', '1', '2', '', '', '', '', '', '', ''],
          ['E5', '10', '2', '2', '0', '', '', '', '', '', '', ''],
          ['E6', '10', '3', '0', '1', '', '', '', '', '', '', ''],
          ['E7', '15', '1', '2', '1', '', '', '', '', '', '', ''],
          ['E8', '15', '2', '0', '2', '', '', '', '', '', '', ''],
          ['E9', '15', '3', '1', '0', '', '', '', '', '', '', ''],
          ['V1', '—', '—', '—', '—', '', '', '', '', '', '', '最优验证'],
          ['V2', '0', '0', '0', '0', '', '', '', '', '', '', '空白对照'],
          ['V3', '10', '0', '0', '0', '', '', '', '', '', '', '基础改性']
        ],
        [600, 700, 700, 700, 800, 700, 700, 700, 700, 900, 800, 1860]
      ),

      // ===== Section 6: Operation Flow =====
      h1('六、操作流程'),

      h2('Day 1-2：材料准备'),
      ...numberedList([
        '采购所有原材料和耗材',
        'PLA 颗粒 60°C 干燥 4 小时',
        'PBAT 颗粒 50°C 干燥 3 小时',
        '按实验分组预称量每组分装到铝箔袋（每袋 200g 总量），贴标签密封'
      ]),

      h2('Day 3-5：挤出实验（E1-E9）'),
      para([txt('每组流程：')]),
      ...numberedList([
        '开机预热料桶至 195°C，稳定 15 分钟',
        '加入对应配方的预混料 200g',
        '螺杆 26 RPM 运转，观察出料状态',
        '牵引机 32 mm/s，冷却水槽常温',
        '绕线，每米标记直径测量点',
        '取 10m 样品用于打印测试',
        '填写实验记录卡',
        '清洗螺杆（空转 PLA 3 分钟），准备下一组'
      ]),

      h2('Day 6-7：数据分析'),
      ...numberedList([
        '汇总 9 组评分到汇总表',
        '正交分析：计算各因素各水平的均值和极差',
        '确定最优组合 → 安排 V1 验证实验',
        '同步进行 V2、V3 对照实验'
      ]),

      h2('Day 8-9：验证与打印测试'),
      ...numberedList([
        '挤出 V1/V2/V3 各 200g',
        '3D 打印机打温度塔（190/200/210°C）',
        '打印光泽样片（10×10×3mm 方块，外墙 3 圈）',
        '光泽度计测量、拍照存档',
        '与市售 silk PLA 样品对比'
      ]),

      h2('Day 10：报告撰写'),
      ...numberedList([
        '正交分析表',
        '最优配方确定',
        '工艺参数推荐',
        '成本核算',
        '下一步优化方向'
      ]),

      // ===== Section 7: Budget =====
      h1('七、预算总览'),
      makeTable(
        ['项目', '金额'],
        [
          ['原材料', '¥765'],
          ['实验耗材', '¥330'],
          ['光泽度计（可选）', '¥300'],
          ['总计', '¥1,395（不含设备）']
        ],
        [4680, 4680]
      ),
      blockquote('如果已有挤出机和 3D 打印机，总投入约 ¥1,400 即可完成全部 12 组实验。'),

      // ===== Section 8: Notes =====
      h1('八、注意事项'),
      ...numberedList([
        '干燥是关键 — PLA 吸潮后挤出会产生气泡，表面起泡完全无法评价光泽。务必充分干燥',
        '珠光粉分散 — 珠光粉容易团聚，建议先用少量 PLA 颗粒和珠光粉在密封袋中摇匀（母粒法），再加入主料',
        '螺杆清洗 — 每组实验之间必须用纯 PLA 过料清洗 3 分钟，避免上一组残留干扰',
        '温度控制 — PBAT 比例 >10% 时可能需要微调温度（±5°C），记录实际温度',
        '安全 — PEG 粉末和珠光粉吸入有害，操作时佩戴口罩和手套',
        '存档 — 每组挤出样品取 1m 保留，用标签贴好实验编号，方便后续对比'
      ]),

      // ===== Section 9: Expected Results =====
      h1('九、预期结果'),
      para([txt('基于文献和行业经验，预期：')]),
      makeTable(
        ['因素', '预期影响'],
        [
          ['PBAT 10-15%', '丝绸光泽最明显；5% 效果弱，15%+ 可能过软'],
          ['PEG 2%', '流动性最佳，光泽提升明显；3%+ 可能太粘'],
          ['珠光粉 1%', '增强丝绸感；2% 可能影响打印性能'],
          ['相容剂 1-2%', 'PBAT ≥10% 时明显改善分散和光泽均匀性']
        ],
        [2400, 6960]
      ),
      spacer(),
      para([
        bold('预测最优配方', { color: COLOR_PRIMARY }),
        txt('：E4 或 E5 附近（PBAT 10%、PEG 2%、珠光粉 1-2%、相容剂 0-2%）')
      ]),
    ]
  }]
});

// ===== Generate =====
const OUTPUT = '/Users/xah/WorkBuddy/2026-06-22-task-1/PLA-Silk-Experiment.docx';
Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync(OUTPUT, buffer);
  console.log('DONE: ' + OUTPUT);
}).catch(err => {
  console.error('ERROR:', err);
  process.exit(1);
});
