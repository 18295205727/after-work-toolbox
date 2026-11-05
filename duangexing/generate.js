const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, LevelFormat, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageBreak
} = require("docx");

// ---------- 样式常量 ----------
const FONT = { ascii: "Times New Roman", hAnsi: "Times New Roman", eastAsia: "Microsoft YaHei" };
const MAIN = "1F4E79";
const DARK = "222222";
const RED  = "C00000";
const HEAD_BG = "D9E2F3";
const LIGHT_BG = "F2F6FC";
const CONTENT_W = 9746;

const border = { style: BorderStyle.SINGLE, size: 4, color: "9DB3D5" };
const borders = { top: border, bottom: border, left: border, right: border };

// ---------- 段落辅助 ----------
function P(text, opts = {}) {
  const runs = Array.isArray(text)
    ? text
    : [new TextRun({ text, bold: !!opts.bold, italics: !!opts.italics, color: opts.color || DARK, size: opts.size || 22 })];
  return new Paragraph({
    spacing: { line: opts.line || 320, before: opts.before || 0, after: opts.after ?? 60 },
    alignment: opts.align || AlignmentType.LEFT,
    indent: opts.indent,
    children: runs,
  });
}

function R(text, o = {}) {
  return new TextRun({ text, bold: !!o.bold, italics: !!o.italics, color: o.color || DARK, size: o.size || 22 });
}

function H1(text) {
  return new Paragraph({
    spacing: { before: 320, after: 140 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: MAIN, space: 4 } },
    children: [new TextRun({ text, bold: true, size: 28, color: MAIN })],
  });
}

function H2(text) {
  return new Paragraph({
    spacing: { before: 200, after: 100 },
    children: [
      new TextRun({ text: "▎", bold: true, size: 24, color: MAIN }),
      new TextRun({ text, bold: true, size: 24, color: MAIN }),
    ],
  });
}

function H3(text) {
  return new Paragraph({
    spacing: { before: 140, after: 60 },
    children: [new TextRun({ text, bold: true, size: 22, color: "2E5C9E" })],
  });
}

function body(text, opts = {}) {
  return P([R(text, opts)], { after: 80 });
}

function bullet(runs, level = 0) {
  const arr = Array.isArray(runs) ? runs : [R(runs)];
  return new Paragraph({
    numbering: { reference: "dots", level },
    spacing: { line: 320, after: 50 },
    children: arr,
  });
}

function cell(children, opts = {}) {
  return new TableCell({
    borders,
    width: { size: opts.width, type: WidthType.DXA },
    shading: opts.fill ? { fill: opts.fill, type: ShadingType.CLEAR, color: "auto" } : undefined,
    verticalAlign: VerticalAlign.CENTER,
    margins: { top: 90, bottom: 90, left: 120, right: 120 },
    children: Array.isArray(children) ? children : [children],
  });
}

function cp(text, opts = {}) {
  return new Paragraph({
    spacing: { line: opts.line || 300, after: opts.after ?? 40 },
    alignment: opts.align || AlignmentType.LEFT,
    children: [new TextRun({ text, bold: !!opts.bold, size: opts.size || 20, color: opts.color || DARK })],
  });
}

function headerRow(labels, widths) {
  return new TableRow({
    tableHeader: true,
    cantSplit: true,
    children: labels.map((t, i) =>
      cell([cp(t, { bold: true, align: AlignmentType.CENTER, size: 21, color: MAIN })],
           { width: widths[i], fill: HEAD_BG })),
  });
}

// ---------- 文档内容 ----------
const children = [];

// 标题
children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { before: 120, after: 60 },
  children: [new TextRun({ text: "《短歌行》教学设计", bold: true, size: 40, color: MAIN })],
}));
children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: 200 },
  children: [new TextRun({
    text: "——「我为曹公正名」诗歌鉴赏项目式学习（共 2 课时）",
    size: 24, color: "555555",
  })],
}));

// 基本信息表
{
  const w = [1500, 3373, 1500, 3373];
  const mk2 = (a, b, c, d) => new TableRow({ cantSplit: true, children: [
    cell([cp(a, { bold: true, color: MAIN, align: AlignmentType.CENTER })], { width: w[0], fill: LIGHT_BG }),
    cell([cp(b)], { width: w[1] }),
    cell([cp(c, { bold: true, color: MAIN, align: AlignmentType.CENTER })], { width: w[2], fill: LIGHT_BG }),
    cell([cp(d)], { width: w[3] }),
  ]});
  children.push(new Table({
    width: { size: 100, type: WidthType.PERCENTAGE }, columnWidths: w,
    rows: [
      mk2("课题", "《短歌行》（曹操）", "教材", "统编版高中语文必修上册 第三单元"),
      mk2("课型", "古典诗词鉴赏·新授课", "课时", "2 课时（每课时 45 分钟）"),
      mk2("授课对象", "高一年级", "教法", "诵读法、文本细读法、范例教学法、情境任务法"),
    ],
  }));
}

// 一、设计理念与大情境
children.push(H1("一、设计理念与核心任务"));
children.push(body("以语文学科核心素养为纲，创设贯穿两课时的真实情境与驱动任务，让学生在「读诗—质疑—解诗—用诗—辩诗」的语言实践活动中完成文本细读。第一课时重「诵读生疑、授人以渔」：教师带学前 3 章，提炼可迁移的语言切入式鉴赏方法；第二课时重「自主探究、成果展评」：学生运用方法独立鉴赏后 5 章，并在「正名」任务中外化学习成果，最终回应单元主题「生命的诗意」——在「人生苦短」的慨叹背后，挺立着曹操求贤若渴、志在一统的英雄生命观。"));

children.push(H2("【大情境】千年「骂名」与一场辩护"));
children.push(body("「宁教我负天下人，休教天下人负我」几乎成了曹操的固定「人设」。但此语出自小说《三国演义》，史载原话为「宁我负人，毋人负我」，语境亦不相同。"));
children.push(body("【总任务】你与同学组成曹操的「谋士辩护团」，以《短歌行》为核心证据，举办一场「诗证·正名」听证会，用诗中的语言文字回答："));
children.push(P([R("诗里的曹操，究竟是一个怎样的人？", { bold: true, color: RED, size: 23 })], { align: AlignmentType.CENTER, after: 120 }));
children.push(body("【成果形式】（小组任选其一，第二课时展示）", { bold: true }));
children.push(bullet([R("A. 辩护词", { bold: true }), R("《为曹公明志》300 字左右，至少引用诗中 3 处语言证据并标明手法。")]));
children.push(bullet([R("B.「诗证据」展板 / 图文推文", { bold: true }), R("：诗句+手法批注+人物结论，可配 AI 生成的意境图。")]));
children.push(bullet([R("C. 60 秒短视频", { bold: true }), R("：AI 数字人「曹操」或 AI 配音朗诵+解说。")]));
children.push(bullet([R("D. 反差人物档案卡", { bold: true }), R("：「世俗印象 VS 诗中真相」对照式卡片。")]));

// 二、教材与学情
children.push(H1("二、教材分析与学情分析"));
children.push(H3("1. 教材分析"));
children.push(body("《短歌行》是建安文学代表作、四言古诗典范。全诗 32 句，每 4 句一章（一「解」），共 8 章；两引《诗经》（《郑风·子衿》《小雅·鹿鸣》），多处运用比兴与用典（朝露、明月、乌鹊、山高海深、周公吐哺），情感在「忧」与「志」之间低回跌宕，是训练「品味诗歌语言、鉴赏表达技巧、把握情感变化」的上佳范本。"));
children.push(H3("2. 学情分析"));
children.push(bullet("学生已接触《静女》等《诗经》作品，对赋比兴有初步印象，但对比兴、用典、引用的系统辨识与规范鉴赏能力尚弱；"));
children.push(bullet("学生对曹操的印象多来自《三国演义》、影视与游戏，刻板印象明显——这恰可转化为探究内驱力；"));
children.push(bullet("高一学生具备诵读与小组合作能力，但独立鉴赏需要清晰的「方法支架」与「任务支架」，故采用「教师示范—学生迁移」的两课时结构。"));

// 三、教学目标
children.push(H1("三、教学目标"));
children.push(bullet([R("语言建构与运用：", { bold: true }), R("准确、有感情地诵读全诗，读准「衿、呦、掇、陌、契、讌、匝、哺」等字音；掌握 8 章基本内容。")]));
children.push(bullet([R("思维发展与提升：", { bold: true }), R("掌握比兴、用典、引用《诗经》的表达技巧；梳理全诗「忧—思—迎—忧—聚—唤—志」的情感变化曲线。")]));
children.push(bullet([R("审美鉴赏与创造：", { bold: true }), R("学会「咬文嚼字→辨识手法→再现情境→体悟情志」的语言切入式鉴赏法，并能迁移方法独立鉴赏后 5 章及同类诗歌。")]));
children.push(bullet([R("文化传承与理解：", { bold: true }), R("理解曹操求贤若渴、建功一统的英雄之志与建安风骨；体认将有限生命融入「天下归心」事业的生命观，辩证、立体地认识历史人物。")]));

// 四、重难点
children.push(H1("四、教学重点与难点"));
children.push(P([R("教学重点：", { bold: true, color: RED }),
  R("掌握诗歌内容；掌握比兴、用典、引用的表达技巧；理解诗人在求贤若渴的「忧」中不乏建功大业的慷慨与英雄之志，把握全诗情感变化。")], { after: 80 }));
children.push(P([R("教学难点：", { bold: true, color: RED }),
  R("从诗歌语言文字切入鉴赏诗歌，掌握可操作的鉴赏角度与方法，并能迁移运用、独立鉴赏。")], { after: 80 }));

// 五、方法与资源
children.push(H1("五、教学方法与资源准备"));
children.push(H3("1. 教学方法"));
children.push(body("诵读涵泳法、问题引导法、文本细读法、范例教学法、小组合作探究法、成果展评法。"));
children.push(H3("2. 资源与技术准备"));
children.push(bullet("名家朗诵音频（如方明、徐涛朗诵《短歌行》）；"));
children.push(bullet("94 版《三国演义》「横槊赋诗」片段（约 2 分钟。提示学生：此为小说影视演绎，非正史记载）；"));
children.push(bullet("《诗经·郑风·子衿》《诗经·小雅·鹿鸣》原文对照打印件；"));
children.push(bullet("AI 工具（按学校条件任选）：对话式 AI 核查字词典故、AI 文生图（「月明星稀，乌鹊南飞」意境图）、AI 配音 / 数字人生成朗诵或辩护视频；"));
children.push(bullet("学生学习单、质疑卡、评价量表（见文末附录，可直接印发）。"));

// ============ 教学过程 ============
children.push(H1("六、教学过程"));

children.push(H2("第一课时：诵读生疑 · 授人以渔（45 分钟）"));
children.push(body("课时目标：读准读熟、提出并合力解决疑难（重点突破引用《诗经》的诗句）；教师示范前 3 章，掌握「语言切入四步法」。", { italics: true, color: "555555" }));

{
  const w = [1560, 760, 3150, 2250, 2026];
  const rows = [headerRow(["教学环节", "时长", "教师活动", "学生活动", "设计意图"], w)];

  const addRow = (name, time, teacher, student, intent) => {
    const paras = arr => arr.map(t => {
      if (typeof t === "string") return cp(t);
      return cp(t.t, { bold: t.bold, color: t.c, size: t.size || 20 });
    });
    rows.push(new TableRow({ cantSplit: true, children: [
      cell([cp(name, { bold: true, color: MAIN, align: AlignmentType.CENTER, size: 20 })], { width: w[0], fill: LIGHT_BG }),
      cell([cp(time, { align: AlignmentType.CENTER })], { width: w[1] }),
      cell(paras(teacher), { width: w[2] }),
      cell(paras(student), { width: w[3] }),
      cell(paras(intent), { width: w[4] }),
    ]}));
  };

  addRow("环节一\n冲突导入\n发布任务", "4′",
    [{t:"播放「横槊赋诗」片段 30 秒，或出示两组标签：左为「奸雄、多疑、宁我负人」（AI 词云），右为「对酒当歌，人生几何」。"},
     {t:"提问：这些标签和这首诗，可能属于同一个人吗？发布「谋士辩护团」总任务。", c:RED}],
    ["观看并说出对曹操的第一印象，产生认知冲突；明确总任务与成果形式。"],
    ["以刻板印象制造悬念，任务驱动学习；点明本课「以诗为据、重新识人」。"]);

  addRow("环节二\n诵读涵泳\n四步读诗", "13′",
    ["① 听读正音：播放名家音频，板书正音——譬(pì)如、子衿(jīn)、呦呦(yōu)、鼓瑟(sè)、可掇(duō)、阡(qiān)陌(mò)、枉(wǎng)、契(qiè)阔、谈讌(yàn)、三匝(zā)、吐哺(bǔ)；",
     "② 领读明节奏：四言以「二二」节拍为主，如「对酒/当歌，人生/几何」；",
     "③ 组织自由读、齐读，提示感慨苍凉的基调；",
     "④ 男女生分章接读，尝试读出「忧」的层次。"],
    ["听音频标注字音、划节拍；", "自由读、齐读、接读；", "在诵读中初步感受语气与情感起伏。"],
    ["以声入文，读准、读顺、读出情；为质疑和鉴赏蓄势。"]);

  addRow("环节三\n质疑问难\n合力解疑", "8′",
    ["要求再读全诗，圈出「看不懂、有疑问」处，写在质疑卡贴上「疑问墙」；",
     {t:"预设难点（尤其引用《诗经》四句）：", c:RED},
     "①「青青子衿，悠悠我心」何意？「君」指谁？",
     "②「呦呦鹿鸣……鼓瑟吹笙」与曹操何干？",
     "③「掇」何义？「乌鹊南飞」写什么？",
     "④「周公吐哺」是什么典故？",
     {t:"组织小组互助+教师点拨：当堂重点解决①②（印发《子衿》《鹿鸣》原文对照），点出「引用」手法；③④存入「问题银行」留待后续。", bold:true}],
    ["质疑、贴卡、小组互助；",
     "对照《诗经》原文发现：",
     "《子衿》写恋人相思→被翻新为对贤才的思慕；",
     "《鹿鸣》写君王宴飨宾客→表达对贤才的盛情礼遇。"],
    ["让真实问题从学生中产生；",
     "引用《诗经》这一难点由师生共同解决；",
     "「问题银行」制造持续期待。"]);

  addRow("环节四\n知人论世\n解题明体", "5′",
    ["解题：「短歌行」为汉乐府旧题（属相和歌平调曲），「行」为古诗体裁，短歌多写宴饮感慨；",
     "背景：曹操，政治家、军事家、诗人，建安文学领袖；诗为宴饮之际抒怀（一说作于建安十三年前后），时天下未定、大业待成；",
     {t:"史料辨析：区分「历史的曹操、《演义》的曹操、诗中的曹操」三个形象。", c:RED}],
    ["圈划课下注释，记录背景；", "讨论：诗中「我」与影视曹操是否同一形象。"],
    ["为理解「忧」与「志」搭建背景支架；", "培养史料意识，呼应总任务。"]);

  addRow("环节五\n精读示范\n习得方法", "12′",
    [{t:"教师带学第 1—3 章，示范「语言切入四步法」（详见下方「精读脚本」）；", bold:true},
     "第 1 章：抓「几何、苦多、当」，辨「譬如朝露」之比兴，悟一「忧」——人生短暂；",
     "第 2 章：设问「何以解忧」，「忧」字重现，借酒写忧；",
     {t:"第 3 章：析「悠悠、沉吟」，明引《子衿》之翻新，悟二「忧」——贤才未至；", c:RED},
     {t:"师生共同提炼并板书：咬文嚼字→辨识手法→再现情境→体悟情志。", bold:true}],
    ["跟随教师圈点批注；", "完成学习单第 1—3 章填空；", "口述四步法的操作要领。"],
    ["范例教学，把鉴赏方法显性化、程序化，为第二课时自主探究搭建支架。"]);

  addRow("环节六\n小结布障\n任务延伸", "3′",
    ["明确 8 章结构（4 句一章）；",
     {t:"布置课后任务：", bold:true},
     "① 用「四步法」独立研读第 4—8 章，完成学习单；",
     "② 4—6 人组成「辩护团」，认领成果形式 A/B/C/D，至少找 3 条「诗证据」；",
     "③ 可用 AI 核查典故，但必须「回到原诗找证据」。"],
    ["组队、分工、认领成果形式；", "记录课后任务与要求。"],
    ["把学习延伸到课外，保障第二课时探究与展示质量。"]);

  children.push(new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, columnWidths: w, rows }));
}

// 精读脚本
children.push(H3("◆ 第 1—3 章精读脚本（教师示范用）"));
children.push(P([R("【第 1 章】", { bold: true, color: MAIN }), R("对酒当歌，人生几何！譬如朝露，去日苦多。")], { after: 30 }));
children.push(bullet([R("咬文嚼字：", { bold: true }), R("「当」（面对）；「几何」「苦多」——感慨什么？")]));
children.push(bullet([R("辨识手法：", { bold: true }), R("「譬如朝露」以朝露易干比喻人生短促（比兴）。")]));
children.push(bullet([R("再现情境：", { bold: true }), R("举杯高歌之际忽生苍凉——乐景写哀，慷慨沉郁。")]));
children.push(bullet([R("体悟情志：", { bold: true }), R("一「忧」——人生短暂。追问：只是贪生畏死吗？设悬。")]));
children.push(P([R("【第 2 章】", { bold: true, color: MAIN }), R("慨当以慷，忧思难忘。何以解忧？唯有杜康。")], { before: 80, after: 30 }));
children.push(bullet("以自问自答的方式直陈「忧思难忘」；借酒消愁，忧思愈深，为全诗定下慷慨悲凉的基调。"));
children.push(P([R("【第 3 章】", { bold: true, color: MAIN }), R("青青子衿，悠悠我心。但为君故，沉吟至今。")], { before: 80, after: 30 }));
children.push(bullet([R("咬文嚼字：", { bold: true }), R("「悠悠」写思念深长不断；「沉吟」指低声吟咏、念念不忘。")]));
children.push(bullet([R("辨识手法：", { bold: true }), R("引用《郑风·子衿》并化用翻新——原诗「纵我不往，子宁不嗣音」写恋人娇嗔相思，此处「君」指贤才，儿女情长化为君臣之思。")]));
children.push(bullet([R("体悟情志：", { bold: true }), R("二「忧」——求贤若渴而贤才未至。")]));

// 第二课时
children.push(new Paragraph({ children: [new PageBreak()] }));
children.push(H2("第二课时：自主探究 · 成果展评（45 分钟）"));
children.push(body("课时目标：迁移「四步法」独立鉴赏第 4—8 章；梳理情感曲线；完成并展评「诗证·正名」成果；升华至单元生命主题。", { italics: true, color: "555555" }));

{
  const w = [1560, 760, 3150, 2250, 2026];
  const rows = [headerRow(["教学环节", "时长", "教师活动", "学生活动", "设计意图"], w)];
  const addRow = (name, time, teacher, student, intent) => {
    const paras = arr => arr.map(t => typeof t === "string"
      ? cp(t)
      : cp(t.t, { bold: t.bold, color: t.c }));
    rows.push(new TableRow({ cantSplit: true, children: [
      cell([cp(name, { bold: true, color: MAIN, align: AlignmentType.CENTER, size: 20 })], { width: w[0], fill: LIGHT_BG }),
      cell([cp(time, { align: AlignmentType.CENTER })], { width: w[1] }),
      cell(paras(teacher), { width: w[2] }),
      cell(paras(student), { width: w[3] }),
      cell(paras(intent), { width: w[4] }),
    ]}));
  };

  addRow("环节一\n温故入新", "3′",
    ["抽查背诵前 3 章；回顾「四步法」；开启「问题银行」剩余问题（掇、乌鹊、周公）。"],
    ["背诵、复述方法；明确本课目标。"],
    ["激活旧知，以旧问题驱动新探究。"]);

  addRow("环节二\n小组探究\n自学后 5 章", "15′",
    ["发放任务单，每组重点研读 1—2 章并填写学习单（教师巡视点拨，支架见下方「第 4—8 章探究支架」）；",
     {t:"同步任务：把找到的「诗证据」贴入本组辩护成果，准备 1 分钟汇报。", c:RED}],
    ["组内分工：朗读、记录、手法分析、成果制作；",
     "按「四步法」独立研读第 4—8 章；",
     "讨论、填表、制作展板/辩护词/视频。"],
    ["方法迁移，在做中学，突破「独立鉴赏」的难点；",
     "为「正名」任务准备语言证据。"]);

  addRow("环节三\n汇报交流\n理清情感曲线", "10′",
    ["组织各组逐章汇报，点拨纠偏；",
     {t:"核心追问：全诗三次写「忧」，曹操「忧」的到底是什么？", c:RED},
     "明确：不忧生死、不忧贫贱——忧人生短暂而功业未就、贤才未集；忧愈深，志愈壮；",
     "板书完成情感曲线，组织齐读尾章。"],
    ["逐章汇报、互相补充；",
     "齐读「山不厌高……天下归心」，读出坚定自信；",
     "记录情感曲线。"],
    ["把握「情感变化」教学重点；",
     "理解「忧中见志」的建安风骨。"]);

  addRow("环节四\n「诗证·正名」听证会", "12′",
    ["主持听证会：各组展示 1 分钟（数字人/配音视频现场播放）；",
     "其余组依量表现场打分；每组展示后接受 30 秒质询：「此结论出自哪个词、哪个典故？」",
     {t:"教师总结：不是为曹操「洗白」，而是以文本为据、立体看人——诗中曹操惜时、敬贤、谦恭、自信、以天下为念。", bold:true}],
    ["展示成果、接受质询、依据量表互评；",
     "记录他组亮点与证据。"],
    ["学习成果显性呈现；",
     "在真实交际情境中落实语言与思维素养。"]);

  addRow("环节五\n升华\n生命的意义", "4′",
    ["补充《短歌行·其二》「不戚年往，忧世不治」与《龟虽寿》「老骥伏枥，志在千里」；",
     {t:"提问：面对「譬如朝露」的有限人生，曹操给出的答案是什么？", c:RED},
     "结语：个体生命短暂，融入「天下归心」的事业便获得重量；读诗，是借诗人的眼睛重新看世界、看他人。"],
    ["联系诗句发言；",
     "记录单元主题札记。"],
    ["回应单元「生命的意义」主题；",
     "形成对曹操的全新认识。"]);

  addRow("环节六\n布置作业", "1′",
    ["① 基础：背诵默写全诗，完成「比兴·用典·引用」考点整理表；",
     "② 提升：诗评《诗里，站着另一个曹操》300 字，或完善辩护词；",
     "③ 拓展（选做）：用 AI 生成意境海报或 30 秒朗诵短视频并附 50 字说明；比较影视曹操与诗中曹操。"],
    ["记录作业，按层选择。"],
    ["分层作业，落实考点并延伸审美与思辨。"]);

  children.push(new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, columnWidths: w, rows }));
}

// 探究支架
children.push(H3("◆ 第 4—8 章探究支架（发学生 / 教师点拨用）"));
const zhiJia = [
  ["第 4 章", "呦呦鹿鸣，食野之苹。我有嘉宾，鼓瑟吹笙。",
    "引用《小雅·鹿鸣》（君王宴飨群臣嘉宾之诗）；以宴饮之礼写对贤才的盛情相迎。", "想象贤才来归的欢欣（虚写/向往）"],
  ["第 5 章", "明明如月，何时可掇？忧从中来，不可断绝。",
    "比兴：以皎皎明月喻贤才高洁而难得（一说喻忧思绵长）；「掇」拾取、摘取，一作「辍」（停止），可作比较探究。", "三「忧」——贤才难致、忧思不绝"],
  ["第 6 章", "越陌度阡，枉用相存。契阔谈讌，心念旧恩。",
    "虚写贤才远道枉驾来访、宴饮叙旧；「枉」字见曹操屈己待士之谦恭。", "欢聚之暖、旧恩之深"],
  ["第 7 章", "月明星稀，乌鹊南飞。绕树三匝，何枝可依？",
    "比兴：以乌鹊徘徊绕树喻贤才流寓四方、择主而事；「三匝」「何枝」含体贴与召唤——天下贤士，何处托身？", "关切、体贴与殷殷召唤"],
  ["第 8 章", "山不厌高，海不厌深。周公吐哺，天下归心。",
    "用典：前二句化用《管子》泰山不弃土石、沧海不拒细流之意，言胸怀广纳；后二句用周公「一沐三握发、一饭三吐哺」之典，以周公自况。卒章显志。", "志——广纳贤才、一统天下"],
];
{
  const w = [900, 2500, 4146, 2200];
  const rows = [headerRow(["章次", "原句", "手法 · 语言要点", "情感"], w)];
  zhiJia.forEach(r => rows.push(new TableRow({ cantSplit: true, children: [
    cell([cp(r[0], { bold: true, color: MAIN, align: AlignmentType.CENTER })], { width: w[0], fill: LIGHT_BG }),
    cell([cp(r[1], { size: 19 })], { width: w[1] }),
    cell([cp(r[2], { size: 19 })], { width: w[2] }),
    cell([cp(r[3], { size: 19 })], { width: w[3] }),
  ]})));
  children.push(new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, columnWidths: w, rows }));
}

// 情感曲线
children.push(H3("◆ 全诗情感曲线（板书同步）"));
children.push(P([R("忧（人生苦短）→ 思（沉吟待贤）→ 迎（鼓瑟吹笙）→ 忧（明月难掇）→ 聚（契阔谈讌）→ 唤（何枝可依）→ 志（天下归心）",
  { bold: true, color: RED, size: 22 })], { align: AlignmentType.CENTER, before: 60, after: 60 }));
children.push(body("「忧」是全诗情感主旋律，「志」是情感归宿：慷慨悲凉其表，英雄之志其里——忧世不治，故求贤若渴；求贤若渴，正为天下归心。"));

// 七、板书设计
children.push(H1("七、板书设计"));
{
  const w = [9746];
  const lines = [
    ["短 歌 行    曹 操", true, MAIN],
    ["（汉乐府旧题 · 四言古诗 · 建安风骨）", false, "555555"],
    ["", false, DARK],
    ["一忧  人生苦短 —— 譬如朝露（比兴）", false, DARK],
    ["二忧  贤才难得 —— 青青子衿 · 呦呦鹿鸣（引用《诗经》）", false, DARK],
    ["三忧  功业未就 —— 明明如月 · 乌鹊南飞（比兴）", false, DARK],
    ["一志  天下归心 —— 山高海深 · 周公吐哺（用典）", true, RED],
    ["", false, DARK],
    ["【鉴赏四步法】咬文嚼字 → 辨识手法 → 再现情境 → 体悟情志", true, MAIN],
    ["慷慨悲凉  ←交织→  求贤若渴 · 建功立业（忧中见志）", false, DARK],
  ];
  const rows = [new TableRow({ cantSplit: true, children: [cell(
    lines.map(([t, b, c]) => new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { line: 340, after: 40 },
      children: [new TextRun({ text: t, bold: b, size: b ? 24 : 21, color: c })],
    })), { width: w[0], fill: LIGHT_BG })] })];
  children.push(new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, columnWidths: w, rows }));
}

// 八、评价量表
children.push(H1("八、学习成果评价量表（听证会现场使用）"));
{
  const w = [1700, 3600, 1100, 1100, 1100, 1146];
  const rows = [headerRow(["评价维度", "评价要点", "优", "良", "中", "得分"], w)];
  const data = [
    ["语言证据（30）", "至少 3 处诗中语言证据，引用准确，能紧扣关键词句（如「悠悠」「沉吟」「枉」）。", "26-30", "21-25", "15-20"],
    ["手法辨识（25）", "比兴、用典、引用判断正确，能说明出处与表达效果（如《子衿》《鹿鸣》、周公吐哺）。", "22-25", "17-21", "12-16"],
    ["情感理解（25）", "准确理解「忧中见志」及情感变化，对曹操形象的结论有文本依据、不脱离诗句。", "22-25", "17-21", "12-16"],
    ["表达与合作（20）", "展示条理清楚、形式恰当；分工明确、回应质询有理有据；AI 使用规范、注明来源。", "17-20", "13-16", "9-12"],
  ];
  data.forEach(d => rows.push(new TableRow({ cantSplit: true, children: [
    cell([cp(d[0], { bold: true, color: MAIN, size: 19, align: AlignmentType.CENTER })], { width: w[0], fill: LIGHT_BG }),
    cell([cp(d[1], { size: 19 })], { width: w[1] }),
    cell([cp(d[2], { align: AlignmentType.CENTER, size: 19 })], { width: w[2] }),
    cell([cp(d[3], { align: AlignmentType.CENTER, size: 19 })], { width: w[3] }),
    cell([cp(d[4], { align: AlignmentType.CENTER, size: 19 })], { width: w[4] }),
    cell([cp("", { align: AlignmentType.CENTER })], { width: w[5] }),
  ]})));
  rows.push(new TableRow({ cantSplit: true, children: [
    cell([cp("合计 / 评委签名", { bold: true, color: MAIN, align: AlignmentType.CENTER, size: 19 })], { width: 4700, fill: LIGHT_BG }),
    cell([cp("", { align: AlignmentType.CENTER })], { width: 3300 }),
    cell([cp("", {})], { width: 1746 }),
  ]}));
  children.push(new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, columnWidths: w, rows }));
}

// 九、作业
children.push(H1("九、分层作业设计"));
children.push(bullet([R("基础类（必做）：", { bold: true }), R("背诵并默写全诗；完成学习单「比兴·用典·引用」考点整理表，标注手法、出处与效果。")]));
children.push(bullet([R("提升类（必做）：", { bold: true }), R("以《诗里，站着另一个曹操》为题写 300 字诗评，观点须有诗句支撑；或完善本组辩护词。")]));
children.push(bullet([R("拓展类（选做）：", { bold: true }), R("① 借助 AI 为《短歌行》生成一张意境海报或 30 秒朗诵短视频，附 50 字创作说明；② 比较《三国演义》中的曹操与《短歌行》中的曹操，写 200 字札记。")]));

// 十、教学反思
children.push(H1("十、教学反思（课后填写）"));
children.push(bullet("诵读环节的时间是否充分？学生是否真正读出了「忧」的层次？"));
children.push(bullet("后 5 章自主探究时，学生能否自觉运用「四步法」？哪些环节仍需支架？"));
children.push(bullet("AI 与音视频资源是否做到了「为学所用、以文为据」？听证会的参与度与生成质量如何？"));
children.push(bullet("学生对曹操「形象反转」的理解是建立在文本证据之上，还是停留在热闹形式？"));

// ============ 附录 ============
children.push(new Paragraph({ children: [new PageBreak()] }));
children.push(H1("附录一  学生学习单：《短歌行》八章鉴赏表"));
children.push(body("班级：________  小组：________  姓名：________     方法提示：咬文嚼字 → 辨识手法 → 再现情境 → 体悟情志", { color: "555555" }));
{
  const w = [760, 2460, 2200, 2126, 2200];
  const rows = [headerRow(["章", "原句", "关键词句", "表达技巧", "情感（一个字/词概括）"], w)];
  const filled = [
    ["1", "对酒当歌，人生几何！譬如朝露，去日苦多。", "几何、苦多；「朝露」", "比兴（以朝露喻人生短促）", "忧——人生短暂（示范）"],
    ["2", "慨当以慷，忧思难忘。何以解忧？唯有杜康。", "忧思难忘、杜康", "设问、借酒写忧", "忧思沉郁（示范）"],
    ["3", "青青子衿，悠悠我心。但为君故，沉吟至今。", "悠悠、沉吟、君", "引用《郑风·子衿》并翻新", "思贤——求贤若渴（示范）"],
  ];
  filled.forEach(d => rows.push(new TableRow({ cantSplit: true, children: [
    cell([cp(d[0], { align: AlignmentType.CENTER, bold: true, color: MAIN })], { width: w[0], fill: "EAF1FB" }),
    cell([cp(d[1], { size: 18 })], { width: w[1] }),
    cell([cp(d[2], { size: 18 })], { width: w[2] }),
    cell([cp(d[3], { size: 18 })], { width: w[3] }),
    cell([cp(d[4], { size: 18 })], { width: w[4] }),
  ]})));
  const blank = [
    ["4", "呦呦鹿鸣，食野之苹。我有嘉宾，鼓瑟吹笙。"],
    ["5", "明明如月，何时可掇？忧从中来，不可断绝。"],
    ["6", "越陌度阡，枉用相存。契阔谈讌，心念旧恩。"],
    ["7", "月明星稀，乌鹊南飞。绕树三匝，何枝可依？"],
    ["8", "山不厌高，海不厌深。周公吐哺，天下归心。"],
  ];
  blank.forEach(d => rows.push(new TableRow({ cantSplit: true, children: [
    cell([cp(d[0], { align: AlignmentType.CENTER, bold: true, color: MAIN })], { width: w[0] }),
    cell([cp(d[1], { size: 18 })], { width: w[1] }),
    cell([cp(" ", {})], { width: w[2] }),
    cell([cp(" ")], { width: w[3] }),
    cell([cp(" ")], { width: w[4] }),
  ]})));
  children.push(new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, columnWidths: w, rows }));
}
children.push(body(""));
children.push(H3("考点整理表：比兴 · 用典 · 引用"));
{
  const w = [1600, 3200, 2623, 2323];
  const rows = [headerRow(["类型", "诗中例句", "出处/本义", "表达效果"], w)];
  const rowsData = [
    ["引用", "青青子衿，悠悠我心", "《诗经·郑风·子衿》，恋人相思", "翻新为对贤才的思慕"],
    ["引用", "呦呦鹿鸣……鼓瑟吹笙", "《诗经·小雅·鹿鸣》，君王宴宾", "表达礼遇贤才的热忱"],
    ["用典", "周公吐哺", "周公「一饭三吐哺」待士", "以周公自况，明示待贤之诚"],
    ["比兴", "譬如朝露", "—", "喻人生短促"],
    ["比兴", "乌鹊南飞，何枝可依", "—", "喻贤才择主、徘徊失所"],
  ];
  rowsData.forEach((d) => {
    rows.push(new TableRow({
      cantSplit: true,
      children: d.map((t, j) =>
        cell(
          [cp(t, { size: 19, bold: j === 0, color: j === 0 ? MAIN : DARK, align: j === 0 ? AlignmentType.CENTER : AlignmentType.LEFT })],
          { width: w[j], fill: j === 0 ? LIGHT_BG : undefined }
        )
      ),
    }));
  });
  for (let i = 0; i < 2; i++) {
    rows.push(new TableRow({ cantSplit: true, children: [
      cell([cp(" ")], { width: w[0] }), cell([cp(" ")], { width: w[1] }),
      cell([cp(" ")], { width: w[2] }), cell([cp(" ")], { width: w[3] }),
    ]}));
  }
  children.push(new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, columnWidths: w, rows }));
}

children.push(H1("附录二  质疑卡（第一课时使用）"));
{
  const w = [9746];
  const inner = [
    new Paragraph({ spacing: { after: 120 }, children: [new TextRun({ text: "我读不懂 / 有疑问的诗句：", bold: true, size: 21, color: MAIN })] }),
    new Paragraph({ spacing: { after: 200 }, border: { bottom: { style: BorderStyle.DASHED, size: 4, color: "999999", space: 6 } }, children: [R(" ")] }),
    new Paragraph({ spacing: { after: 120 }, children: [new TextRun({ text: "我的初步理解或猜测：", bold: true, size: 21, color: MAIN })] }),
    new Paragraph({ spacing: { after: 200 }, border: { bottom: { style: BorderStyle.DASHED, size: 4, color: "999999", space: 6 } }, children: [R(" ")] }),
    new Paragraph({ spacing: { after: 60 }, children: [new TextRun({ text: "讨论后我们的结论（运用了什么方法/资料）：", bold: true, size: 21, color: MAIN })] }),
    new Paragraph({ border: { bottom: { style: BorderStyle.DASHED, size: 4, color: "999999", space: 6 } }, children: [R(" ")] }),
  ];
  const rows = [new TableRow({ cantSplit: true, children: [cell(inner, { width: w[0] })] })];
  children.push(new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, columnWidths: w, rows }));
}

children.push(H1("附录三  「诗证·正名」辩护任务单"));
children.push(bullet([R("我们的辩护对象：", { bold: true }), R("为曹操摘掉「宁教我负天下人，休教天下人负我」的标签。")]));
children.push(bullet([R("我们的核心主张（一句话）：", { bold: true }), R("诗中的曹操是一位 ________________________________ 的人。")]));
children.push(bullet([R("诗证据清单（≥3 条）：", { bold: true })]));
children.push(body("证据一：诗句「________________」——关键词/手法：________——证明：________________", { indent: { left: 360 } }));
children.push(body("证据二：诗句「________________」——关键词/手法：________——证明：________________", { indent: { left: 360 } }));
children.push(body("证据三：诗句「________________」——关键词/手法：________——证明：________________", { indent: { left: 360 } }));
children.push(bullet([R("我们选择的成果形式：", { bold: true }), R("□ A 辩护词   □ B 展板/推文   □ C 数字人短视频   □ D 反差人物档案卡")]));
children.push(bullet([R("分工：", { bold: true }), R("朗读____ 记录____ 分析____ 制作/汇报____")]));
children.push(H3("AI 使用提示词范例与三条原则"));
children.push(body("提示词 1：请帮我查证「周公吐哺」的典故出处与原意，注明文献名称，我会自己核对原文。", { indent: { left: 360 } }));
children.push(body("提示词 2：「青青子衿」出自《诗经》哪一篇？原诗表达什么情感？曹操借用后意思发生了什么变化？", { indent: { left: 360 } }));
children.push(body("提示词 3：请生成一幅「月明星稀、乌鹊南飞、江水畔宴饮」的中国水墨画风格意境图，横版，构图留白。", { indent: { left: 360 } }));
children.push(bullet([R("三原则：", { bold: true, color: RED }), R("① 先读诗、再问 AI；② AI 的答案必须回到原诗与可靠文献中验证（记录一条「AI 也会错」可加分）；③ AI 是助手，成果由小组署名并负责。")]));

// ---------- 生成 ----------
const doc = new Document({
  styles: { default: { document: { run: { font: FONT, size: 22, color: DARK } } } },
  numbering: {
    config: [{
      reference: "dots",
      levels: [
        { level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 480, hanging: 260 } } } },
        { level: 1, format: LevelFormat.BULLET, text: "◦", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 840, hanging: 260 } } } },
      ],
    }],
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 },
      },
    },
    children,
  }],
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("/workspace/短歌行教学设计.docx", buffer);
  console.log("done", buffer.length);
});
