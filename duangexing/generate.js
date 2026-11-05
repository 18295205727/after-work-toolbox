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
const GREEN = "385723";
const HEAD_BG = "D9E2F3";
const LIGHT_BG = "F2F6FC";
const WARN_BG = "FBEFE3";

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
function body(text, opts = {}) { return P([R(text, opts)], { after: 80 }); }
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
    children: [new TextRun({ text, bold: !!opts.bold, italics: !!opts.italics, size: opts.size || 20, color: opts.color || DARK })],
  });
}
function headerRow(labels, widths) {
  return new TableRow({
    tableHeader: true, cantSplit: true,
    children: labels.map((t, i) =>
      cell([cp(t, { bold: true, align: AlignmentType.CENTER, size: 21, color: MAIN })],
           { width: widths[i], fill: HEAD_BG })),
  });
}
function makeTable(w, rows) {
  return new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, columnWidths: w, rows });
}
function txt(t) { return typeof t === "string" ? cp(t) : cp(t.t, { bold: t.bold, color: t.c, italics: t.i, size: t.size || 20 }); }

// ================= 文档内容 =================
const children = [];

// 标题
children.push(new Paragraph({
  alignment: AlignmentType.CENTER, spacing: { before: 120, after: 60 },
  children: [new TextRun({ text: "《短歌行》教学设计", bold: true, size: 40, color: MAIN })],
}));
children.push(new Paragraph({
  alignment: AlignmentType.CENTER, spacing: { after: 180 },
  children: [new TextRun({ text: "——「我为曹公正名」诗歌鉴赏项目式学习（共 2 课时 · 教师精讲前 4 章·学生探究后 4 章）", size: 22, color: "555555" })],
}));

// 基本信息
{
  const w = [1500, 3373, 1500, 3373];
  const mk2 = (a, b, c, d) => new TableRow({ cantSplit: true, children: [
    cell([cp(a, { bold: true, color: MAIN, align: AlignmentType.CENTER })], { width: w[0], fill: LIGHT_BG }),
    cell([cp(b)], { width: w[1] }),
    cell([cp(c, { bold: true, color: MAIN, align: AlignmentType.CENTER })], { width: w[2], fill: LIGHT_BG }),
    cell([cp(d)], { width: w[3] }),
  ]});
  children.push(makeTable(w, [
    mk2("课题", "《短歌行》（曹操）", "教材", "统编版高中语文必修上册 第三单元"),
    mk2("课型", "古典诗词鉴赏·新授课", "课时", "2 课时（每课时 45 分钟）"),
    mk2("授课对象", "高一年级", "教法", "诵读法、范例教学法、文本细读、合作探究、即席表达"),
  ]));
}

// 一、设计理念
children.push(H1("一、设计理念与核心任务"));
children.push(body("以语文学科核心素养为纲，用一个贯穿两课时的真实情境驱动学习。教师精讲前 4 章，把「从语言切入鉴赏诗歌」的方法教透、把支点给足；学生运用方法自主探究后 4 章，教师点拨收束，最终在当堂「诗证·正名」听证中外化成果，并以一篇文学短评收束全课、对接单元写作任务。全课强调：读出的不是一个沉溺「我颓」的曹操，而是一个慨叹时光、渴求英才、矢志一统的「我要」的曹操。"));

children.push(H2("【大情境】千年「骂名」与一场当堂辩护"));
children.push(body("「宁教我负天下人，休教天下人负我」几乎成了曹操的固定「人设」。但此语出自小说《三国演义》，史载原话为「宁我负人，毋人负我」，语境亦不相同。"));
children.push(body("【总任务】假如你是曹操的谋士，本诗就是主公的当庭自述。第二课时将举行一场当堂「诗证·正名」听证——不借助任何打印材料，仅凭课本、笔记本与课堂生成的批注，即席回答："));
children.push(P([R("诗里的曹操，究竟是一个怎样的人？", { bold: true, color: RED, size: 23 })], { align: AlignmentType.CENTER, after: 100 }));
children.push(body("【成果呈现（当堂完成，无需打印）】", { bold: true }));
children.push(bullet([R("课中：", { bold: true }), R("每人在笔记本上写「一句主张 + 三条诗证据」，小组即席陈词 1 分钟并接受质询，全班星级评价。")]));
children.push(bullet([R("课后：", { bold: true }), R("一篇 300 字左右文学短评（对接单元写作任务），并可选做 AI 意境图 / 数字人朗诵短视频。")]));

// 二、教材学情
children.push(H1("二、教材分析与学情分析"));
children.push(H3("1. 教材分析"));
children.push(body("《短歌行》是建安文学代表作、四言古诗典范。全诗 32 句，每 4 句一章，共 8 章；两引《诗经》（《郑风·子衿》《小雅·鹿鸣》），多处运用比兴与用典（朝露、明月、乌鹊、山高海深、周公吐哺）。诗中「忧」与「志」回环激荡：以男女之情写求贤之思，以宴飨之礼写待贤之诚，以圣人之典写一统之志——是训练「品味语言、鉴赏手法、把握情感变化」并理解「手法如何服务于政治理想」的上佳范本。"));
children.push(H3("2. 学情分析"));
children.push(bullet("学生已接触《诗经》赋比兴，但对比兴、用典、引用的辨识及「手法→情志」的规范分析尚弱，需要可操作的方法支点；"));
children.push(bullet("学生对曹操的印象多来自《三国演义》、影视游戏，刻板印象明显——可转化为探究内驱力；"));
children.push(bullet("学生易把「人生几何」读成及时行乐、把「忧」读成消沉颓丧——诵读时须以明确支点引导其读出「忧而不颓、悲中见壮」；"));
children.push(bullet("「还原情境」对学生最难（距今年代久远），须以「五要素支架 + 可直接生成画面的情境表述」降低难度，并自然衔接 AI 技术运用。"));

// 三、目标
children.push(H1("三、教学目标"));
children.push(bullet([R("语言建构与运用：", { bold: true }), R("读准「衿、呦、掇、陌、枉、契、讌、匝、哺」等字音，依「二二」节拍有感情地诵读；掌握 8 章基本内容并背诵全诗。")]));
children.push(bullet([R("思维发展与提升：", { bold: true }), R("掌握比兴、用典、引用的表达技巧，能具体分析每处手法「如何」传达曹操的政治理想与情感世界；梳理「三忧一志」的情感曲线。")]));
children.push(bullet([R("审美鉴赏与创造：", { bold: true }), R("学会「咬文嚼字→辨识手法→还原情境→体悟情志」四步鉴赏法，重点掌握「还原情境五要素」，并独立迁移到后 4 章及同类诗歌。")]));
children.push(bullet([R("文化传承与理解：", { bold: true }), R("读出曹操惜时如金、求贤若渴、以天下为念的英雄生命观；对读《求贤令》，立体认识历史人物，理解单元「生命的意义」主题。")]));

// 四、重难点
children.push(H1("四、教学重点与难点"));
children.push(P([R("教学重点：", { bold: true, color: RED }),
  R("掌握诗歌内容与比兴、用典、引用的技巧，并说清这些手法如何精准传达曹操的政治理想与情感世界；把握「忧中见志」的情感变化。")], { after: 80 }));
children.push(P([R("教学难点：", { bold: true, color: RED }),
  R("① 诵读中以声传情，读出「我要」的英雄之气而非「我颓」的消沉；② 凭语言支点（尤其「还原情境五要素」）独立鉴赏后 4 章，并在无打印材料的条件下当堂完成有理有据的即席表达。")], { after: 80 }));

// 五、方法资源
children.push(H1("五、教学方法与资源准备"));
children.push(H3("1. 教学方法"));
children.push(body("诵读涵泳法（支架导读）、范例教学法（前 4 章精讲示范）、小组合作探究法（后 4 章）、情境任务法（当堂听证）、读写结合法（文学短评）。"));
children.push(H3("2. 资源与技术（全部以投影 / 板书呈现，学生无需打印材料）"));
children.push(bullet("名家朗诵音频（方明 / 徐涛版）、94 版《三国演义》「横槊赋诗」片段约 30 秒（注明：小说演绎，非正史）；"));
children.push(bullet("《诗经·子衿》《鹿鸣》原文、《求贤令》节选——制成投影幻灯片；"));
children.push(bullet("AI 工具（按校情选用）：对话式 AI 核典故、AI 文生图当堂验证「还原情境」、AI 配音 / 数字人生成朗诵短视频；"));
children.push(bullet("学生只需：课本、笔记本、练习本；教师准备：便利贴（疑问墙）、三星评价牌（红 / 黄 / 绿卡或手势）。"));

// ================= 六、教学过程 =================
children.push(H1("六、教学过程"));

// ---------- 第一课时 ----------
children.push(H2("第一课时：诵读立骨 · 精讲授法（45 分钟）"));
children.push(body("课时目标：读出「忧而不颓」的英雄之气；提出并合力解决疑难（重点：两引《诗经》）；教师精讲第 1—4 章，教透「四步鉴赏法」与「还原情境五要素」。", { italics: true, color: "555555" }));
{
  const w = [1560, 720, 3190, 2250, 2026];
  const rows = [headerRow(["教学环节", "时长", "教师活动", "学生活动", "设计意图"], w)];
  const addRow = (name, time, teacher, student, intent) => {
    rows.push(new TableRow({ cantSplit: true, children: [
      cell([cp(name, { bold: true, color: MAIN, align: AlignmentType.CENTER, size: 20 })], { width: w[0], fill: LIGHT_BG }),
      cell([cp(time, { align: AlignmentType.CENTER })], { width: w[1] }),
      cell(teacher.map(txt), { width: w[2] }),
      cell(student.map(txt), { width: w[3] }),
      cell(intent.map(txt), { width: w[4] }),
    ]}));
  };

  addRow("环节一\n冲突入课", "3′",
    [{t:"播放「横槊赋诗」30 秒，板书两个词：「奸雄」与「对酒当歌」。"},
     {t:"提问：世人都说他「宁我负人」，可他自己的诗里这样说吗？发布总任务——下节课当堂「诗证·正名」，今天先学会如何从诗句中找证据。", c:RED}],
    ["说对曹操的第一印象；", "明确任务与两课节奏：今天学方法，明天我来辩。"],
    ["以刻板印象制造认知冲突；", "任务预告让学法有目的感。"]);

  addRow("环节二\n诵读涵泳\n支架导读", "14′",
    [{t:"① 听读正音（3′）：听名家音频，投影正音与「二二」节拍；", },
     "② 对比辨调（3′）：教师用「颓读法」（叹气、拖音、萎靡）与「英雄读法」（沉郁、饱满、尾句上扬）各读第 1 章，问：哪一种才是曹操？",
     {t:"③ 依「诵读支点表」领读（5′，表见下方）：逐章标重音、语速、神情，强调「忧字往心里沉，志字向远方放」；", bold:true},
     "④ 练读展示（3′）：自由读→同桌互听（看对方眼神是否「抬起来」）→男女生分声部接读 8 章。"],
    ["标音、划拍；",
     {t:"在对比中辨析：哀而不伤、忧而不颓才是本诗基调；"},
     "跟读、练读、接读；",
     "读尾章时声音渐强、目光上扬。"],
    ["诵读给「支点」而非空喊「有感情」；",
     "在细腻的声气处理中把时光之叹与求贤之渴融为一体，落实「我要」精神。"]);

  addRow("环节三\n疑问墙\n合力解疑", "7′",
    ["请学生再读全诗，把不懂处写在便利贴上，贴上黑板「疑问墙」；",
     {t:"重点当堂共解（投影《子衿》《鹿鸣》原文）：①「青青子衿」何意、「君」指谁；②「呦呦鹿鸣」为何写宴饮；③「掇」何义；④「周公吐哺」何典。", c:RED},
     "①②当堂解透并点出「引用」；③④存入「问题银行」，下节课学生自主解锁。"],
    ["质疑、贴墙；",
     "对照《诗经》原文发现：恋歌→思贤，宴宾→迎贤；",
     "记录未解问题。"],
    ["让真问题由学生产生；",
     "引用《诗经》的难点师生共解；问题银行形成课际悬念。"]);

  addRow("环节四\n解题知人", "4′",
    ["解题：汉乐府旧题，「行」为体裁；",
     "背景：曹操集政治家、军事家、诗人于一身，建安文学领袖；诗作于天下未定、大业待成之际（一说建安十三年前后）；",
     {t:"立三个曹操：历史的曹操 /《演义》的曹操 / 诗中的曹操——本课只读第三个。", bold:true}],
    ["圈划注释，记录背景；", "齐读作者简介。"],
    ["为情志理解搭背景支架；", "以史料意识防止用影视形象解诗。"]);

  addRow("环节五\n精讲第 1—4 章\n示范四步法", "15′",
    [{t:"按「精读脚本」（见下方）精讲：第 1 章完整走一遍四步法，重点演示「还原情境五要素」并生成「AI 情境表述」；第 2 章扶着走；第 3、4 章重点析两引《诗经》，追问「引用如何替曹操说出政治理想」；", bold:true},
     {t:"板书提炼方法：咬文嚼字→辨识手法→还原情境（时/地/人/事/态）→体悟情志。", c:RED}],
    ["圈点批注，填笔记本上的示范栏；",
     "口述「还原情境」五要素；",
     "用一句话说出两引《诗经》写出了怎样的待贤之心。"],
    ["范例教学：把鉴赏方法与「手法服务于理想」的分析路径显性化、程序化。"]);

  addRow("环节六\n小结布学", "2′",
    ["回顾 8 章结构与四步法；",
     "课后：① 熟读成诵前 4 章；② 用四步法在笔记本批注第 5—8 章，每章至少写出「一个词 + 一种手法 + 一句情境话」；③ 有条件者用附录提示词让 AI 生成第 7 章画面，下节课带来验证。"],
    ["记录任务，明确「带证据上课」。"],
    ["保障第二课时当堂探究与即席听证的质量。"]);

  children.push(makeTable(w, rows));
}

// 诵读支点表
children.push(H3("◆ 诵读支点表（投影 · 教师领读用）"));
children.push(body("总要求：气息饱满，基调「慷慨悲凉、忧而不颓」——声音可以沉下去，精神必须立起来；尾章一章比一章亮。", { bold: true, color: MAIN }));
{
  const w = [820, 2620, 1700, 2580, 2026];
  const rows = [headerRow(["章", "诗句（首句）", "情感", "重音 · 语速", "神情 / 形体提示"], w)];
  const data = [
    ["1", "对酒当歌，人生几何", "苍凉一问", "重音「几何」「苦多」；语速放缓", "不叹气、不摇头；目光平视远方——是逼问时间，非畏惧时间"],
    ["2", "慨当以慷，忧思难忘", "忧思沉郁", "「难忘」延长；「杜康」轻轻一转", "眉头微蹙随即舒展，自嘲而不自弃"],
    ["3", "青青子衿，悠悠我心", "深长思念", "「悠悠」二字轻而绵长；「沉吟」渐强", "如低声自语，手抚胸口——思念是有温度的"],
    ["4", "呦呦鹿鸣，鼓瑟吹笙", "明朗期盼", "节奏加快，重音「嘉宾」", "面色转暖、略带笑意——想象贤才临门的欢喜"],
    ["5", "明明如月，何时可掇", "仰望追问", "「何时可掇」抬头一问；「不可断绝」沉实", "抬头望月，语气坚定：求之不得，亦不停歇"],
    ["6", "越陌度阡，心念旧恩", "温暖恳切", "「枉」字轻读表敬意", "如迎客上门，身体微前倾，谦逊有礼"],
    ["7", "月明星稀，何枝可依", "体贴召唤", "景语速缓；「何枝可依」放软声气", "不是质问人才为何观望，而是替他担忧、向他招手"],
    ["8", "山不厌高，天下归心", "豪迈誓志", "逐句渐强；「归心」收束有力、余韵上扬", "挺身而立、目光望远——读出「我要」的雄心与自信"],
  ];
  data.forEach(d => rows.push(new TableRow({ cantSplit: true, children: [
    cell([cp(d[0], { bold: true, color: MAIN, align: AlignmentType.CENTER })], { width: w[0], fill: LIGHT_BG }),
    cell([cp(d[1], { size: 19 })], { width: w[1] }),
    cell([cp(d[2], { size: 19 })], { width: w[2] }),
    cell([cp(d[3], { size: 19 })], { width: w[3] }),
    cell([cp(d[4], { size: 19 })], { width: w[4] }),
  ]})));
  children.push(makeTable(w, rows));
}

// 精读脚本
children.push(H3("◆ 第 1—4 章精读脚本（精讲示范用）"));

children.push(P([R("【第 1 章｜完整示范四步法】", { bold: true, color: MAIN }), R("对酒当歌，人生几何！譬如朝露，去日苦多。")], { after: 30 }));
children.push(bullet([R("第一步 · 咬文嚼字：", { bold: true }), R("「当」，面对；「几何」，（人生）能有多少；「苦多」，可恨逝去的日子太多。")]));
children.push(bullet([R("第二步 · 辨识手法：", { bold: true }), R("比兴——以「朝露」见日即干起兴作比。")]));
children.push(bullet([R("第三步 · 还原情境（五要素支架）：", { bold: true, color: RED }),
  R("何时——建安年间的一个秋夜；何地——江畔中军大帐的酒宴；何人——年逾五旬的曹操与满座将佐；何事——酒酣乐作之际忽然停杯；何态——神色由欢转肃，望月不语。")]));
children.push(P([R("　　合成「情境一句话」（此句即后期 AI 生成画面 / 视频的标准表述）：", { bold: true })], { after: 30 }));
children.push(new Paragraph({
  spacing: { line: 320, after: 80 }, indent: { left: 360 },
  shading: { fill: WARN_BG, type: ShadingType.CLEAR, color: "auto" },
  border: { left: { style: BorderStyle.SINGLE, size: 18, color: "C55A11", space: 8 } },
  children: [new TextRun({ text: "东汉末年，天下未定；月明星稀的秋夜，江畔军营夜宴正酣，鼓乐声中，年逾五旬的曹操忽然停杯仰望明月，看见草叶上的朝露转瞬将晞，神色苍凉而目光坚定。", size: 21 })],
}));
children.push(bullet([R("第四步 · 体悟情志：", { bold: true }),
  R("一「忧」——人生短暂。但紧追一问：他是怕死吗？不。朝露之喻逼出的是时不我待——正因来日无多，愈要抓紧平定天下。「颓」读是及时行乐，「壮」读才是英雄本色。", )]));

children.push(P([R("【第 2 章｜扶着走】", { bold: true, color: MAIN }), R("慨当以慷，忧思难忘。何以解忧？唯有杜康。")], { before: 90, after: 30 }));
children.push(bullet("自问自答直陈「忧思难忘」。追问：酒真能解忧吗？——酒只是话头，忧思越浇越盛，为下文揭出「忧的内容」蓄势。诵读要点：「杜康」二字是自嘲的一转，不可读成醉生梦死。"));

children.push(P([R("【第 3、4 章｜重点突破：两引《诗经》怎样说出政治理想】", { bold: true, color: MAIN })], { before: 90, after: 30 }));
children.push(bullet([R("第 3 章 ", { bold: true }), R("青青子衿，悠悠我心。但为君故，沉吟至今。", { bold: true })]));
children.push(bullet([R("出处：", { bold: true }), R("《郑风·子衿》本是女子思念恋人之歌（「纵我不往，子宁不嗣音」）。曹操只取其深情，舍去娇嗔——「君」，正是天下贤才。", )], 1));
children.push(bullet([R("为何精准：", { bold: true }), R("一纸求贤令可以写「快来投奔」，诗却把招聘写成了朝思暮想。以爱情之深喻求贤之渴，政治号召变成了情感呼唤——见出他求贤不是权术，而是真心。", )], 1));
children.push(bullet([R("第 4 章 ", { bold: true }), R("呦呦鹿鸣，食野之苹。我有嘉宾，鼓瑟吹笙。", { bold: true })]));
children.push(bullet([R("出处：", { bold: true }), R("《小雅·鹿鸣》是周代君王宴飨群臣嘉宾的礼乐之歌。", )], 1));
children.push(bullet([R("为何精准：", { bold: true }), R("引此诗，曹操便自置于「礼贤之主」的位置，向天下人承诺：你来，便是我的「嘉宾」，我以礼乐、以国事相待。这是写欢宴，更是发表政治姿态——愿与贤才共治天下。", )], 1));
children.push(bullet([R("两章合看：", { bold: true, color: RED }), R("第 3 章是「盼你来」的私语，第 4 章是「敬你如宾」的公示——由思而迎，情感升温；一个内心柔软、姿态谦恭却志在天下的政治家形象，全在借来的八句诗里立了起来。")]));

// ---------- 第二课时 ----------
children.push(new Paragraph({ children: [new PageBreak()] }));
children.push(H2("第二课时：探究后 4 章 · 当堂听证（45 分钟）"));
children.push(body("课时目标：运用四步法自主探究第 5—8 章（教师点拨）；完成手法与情感总结；在无打印材料的条件下完成「一句主张 + 三条诗证据」的当堂听证；对读《求贤令》、发布文学短评任务。", { italics: true, color: "555555" }));
{
  const w = [1560, 720, 3190, 2250, 2026];
  const rows = [headerRow(["教学环节", "时长", "教师活动", "学生活动", "设计意图"], w)];
  const addRow = (name, time, teacher, student, intent) => {
    rows.push(new TableRow({ cantSplit: true, children: [
      cell([cp(name, { bold: true, color: MAIN, align: AlignmentType.CENTER, size: 20 })], { width: w[0], fill: LIGHT_BG }),
      cell([cp(time, { align: AlignmentType.CENTER })], { width: w[1] }),
      cell(teacher.map(txt), { width: w[2] }),
      cell(student.map(txt), { width: w[3] }),
      cell(intent.map(txt), { width: w[4] }),
    ]}));
  };

  addRow("环节一\n温故开箱", "3′",
    ["齐背前 4 章；回顾四步法与「还原情境五要素」；",
     {t:"开启「问题银行」：第 5 章「掇」、第 7 章「乌鹊」、第 8 章「周公」三题，今天由你们自己解锁。", c:RED}],
    ["齐背、复述方法；", "认领探究问题。"],
    ["激活旧知，以旧问题驱动新探究。"]);

  addRow("环节二\n自主探究\n第 5—8 章", "14′",
    [{t:"分组（4 组各主研 1 章），投影「探究四问」支架：", bold:true},
     "① 章眼是哪一个词/句？② 用了什么手法？③ 用五要素写一句「情境话」（有平板的组可现场输入 AI 生图，用画面检验情境是否合理）；④ 此章可见曹操怎样的情志？",
     "教师巡视点拨：只追问、不代答（点拨要点见下方支架表）；",
     "要求每生在笔记本留下「四步批注」，并圈出 1 条最有力的「诗证据」。"],
    ["组内分工：朗读、记录、手法分析、汇报；",
     "按四问独立研读、写批注；",
     "可选：现场生成 AI 意境图并据图修正自己的情境描述；",
     "圈定诗证据。"],
    ["方法迁移、在做中学，突破独立鉴赏难点；",
     "AI 用于检验「还原情境」，而非替代思考；",
     "为当堂听证储备证据。"]);

  addRow("环节三\n汇报明理\n手法情感总结", "8′",
    ["各组逐章汇报（每组 1.5 分钟），教师相机点拨，重点追问：这处手法替曹操说出了什么政治理想？",
     {t:"汇报后完成两项总结（见下方专节）：手法总结——三「借」；情感总结——三忧一志。", bold:true},
     "全班齐读尾章，要求读出「归心」的力量。"],
    ["汇报、互评、补注；",
     "记录手法与情感总结；",
     "齐读：沉郁起、豪壮收。"],
    ["把零散发现上升为结构化认识；",
     "落实「手法服务于政治理想」与情感变化的重点。"]);

  addRow("环节四\n当堂听证\n「诗证·正名」", "12′",
    [{t:"全程只用课本与笔记本，流程投影：", bold:true},
     "① 发布辩题（0.5′）：为曹操摘掉「宁我负人」标签；",
     "② 当堂起草（3′）：每人在练习本写「一句主张+三条诗证据」（模板投影，见附录二）；",
     "③ 组内互证（2′）：两人轮流即席陈述，对方只追问「这出自哪个词、哪个典？」；",
     "④ 听证陈词（5.5′）：每组抽 1 人即席陈词 1 分钟+接受 30 秒质询；",
     "⑤ 全班举牌：证据来自诗句★、手法判断准确★、结论令人信服★。",
     {t:"教师收束：不是「洗白」，而是以诗为据、立体看人。", c:RED}],
    ["当堂写主张、列证据；",
     "即席陈词、互相质询；",
     "依三星标准举牌评价；",
     "记录最有说服力的一条证据。"],
    ["任务完全当堂完成，不依赖打印；",
     "在真实言语实践中呈现学习成果，落实语言与思维素养。"]);

  addRow("环节五\n升华对读\n生命的意义", "5′",
    ["投影《求贤令》节选（「唯才是举，吾得而用之」），与「嘉宾」「吐哺」对读印证；",
     "再联《短歌行·其二》「不戚年往，忧世不治」、《龟虽寿》「老骥伏枥，志在千里」；",
     {t:"一问收束：面对「譬如朝露」的有限人生，曹操的答案是什么？", c:RED}],
    ["对读发言：诗是「求贤」的深情版，令是「求贤」的政令版；",
     "写一句单元主题札记：有限的生命怎样获得重量。"],
    ["诗史互证，理解曹操一以贯之的人才观；",
     "回应单元「生命的意义」主题。"]);

  addRow("环节六\n布置写作", "3′",
    [{t:"发布文学短评任务（对接本单元写作）：", bold:true},
     "四选一题目，300 字左右，要求「一个口子、三处诗证、叙议结合」（支架见附录三）；",
     "选做：用附录四提示词生成意境图或数字人朗诵短视频，附 50 字创作说明。"],
    ["记录题目与要求；", "课堂口头拟一个观点句。"],
    ["由读到写，完成鉴赏方法的迁移固化。"]);

  children.push(makeTable(w, rows));
}

// 第5-8章点拨支架
children.push(H3("◆ 第 5—8 章探究支架（教师点拨用 · 投影）"));
{
  const w = [820, 2400, 4306, 2220];
  const rows = [headerRow(["章", "原句", "点拨要点（只在学生卡壳时追问）", "情志"], w)];
  const data = [
    ["5", "明明如月，何时可掇？忧从中来，不可断绝。",
      "「掇」，拾取、摘取（一作「辍」，停止，可比较：哪个字更好？）；以皎皎明月喻什么？——贤才高洁可望而不可即，也喻理想皎洁。",
      "三「忧」：贤才难致而求索不绝"],
    ["6", "越陌度阡，枉用相存。契阔谈讌，心念旧恩。",
      "这是眼前实景还是想象？（虚写）「枉」字为何精妙？——贤者屈尊来访，一个「枉」字写出曹操屈己待人之谦。",
      "欢聚之暖、念旧之诚"],
    ["7", "月明星稀，乌鹊南飞。绕树三匝，何枝可依？",
      "乌鹊喻谁？「三匝」「何枝」是谁的口气？——不是指责贤才观望犹豫，而是替他焦虑栖身之所：体贴之中含召唤。",
      "对天下游子的体贴与召唤"],
    ["8", "山不厌高，海不厌深。周公吐哺，天下归心。",
      "前二句化用《管子》言海纳百川；「周公吐哺」用周公「一饭三吐哺、一沐三握发」之典——曹操以谁自况？向天下许下什么承诺？",
      "卒章显志：广纳贤才、一统天下"],
  ];
  data.forEach(d => rows.push(new TableRow({ cantSplit: true, children: [
    cell([cp(d[0], { bold: true, color: MAIN, align: AlignmentType.CENTER })], { width: w[0], fill: LIGHT_BG }),
    cell([cp(d[1], { size: 19 })], { width: w[1] }),
    cell([cp(d[2], { size: 19 })], { width: w[2] }),
    cell([cp(d[3], { size: 19 })], { width: w[3] }),
  ]})));
  children.push(makeTable(w, rows));
}

// ============ 七、手法与情感总结 ============
children.push(H1("七、手法总结与情感总结（第二课时板书生成）"));
children.push(H3("1. 手法总结：三「借」——每一处手法都在替他「言志」"));
{
  const w = [1200, 2560, 3060, 2926];
  const rows = [headerRow(["手法", "诗中例句", "出处 / 喻义", "精准传达的政治理想与情感"], w)];
  const data = [
    ["引用（借民歌）", "青青子衿，悠悠我心", "《郑风·子衿》：恋歌", "把求贤写成深情相思——见求贤之真、之柔"],
    ["引用（借礼乐）", "呦呦鹿鸣……鼓瑟吹笙", "《小雅·鹿鸣》：君王宴宾", "以礼贤之主自任，许诺与贤才共治天下"],
    ["比兴（借物象）", "譬如朝露", "朝露易晞", "人生短促之叹→时不我待、抓紧建功的紧迫感"],
    ["比兴（借物象）", "明明如月，何时可掇", "明月高洁难取", "以高洁许贤才，写尽仰慕与求索不止"],
    ["比兴（借物象）", "乌鹊南飞，何枝可依", "乌鹊择枝", "不责贤者观望，反替其择主焦虑——包容的政治胸怀"],
    ["用典（借圣君）", "山不厌高，海不厌深", "化用《管子》", "执政者的容量：唯才是举、不拘一格"],
    ["用典（借圣君）", "周公吐哺，天下归心", "周公一饭三吐哺", "以周公自况，公开许下待贤治国、一统天下的政治承诺"],
  ];
  data.forEach(d => rows.push(new TableRow({ cantSplit: true, children: [
    cell([cp(d[0], { bold: true, color: MAIN, align: AlignmentType.CENTER, size: 19 })], { width: w[0], fill: LIGHT_BG }),
    cell([cp(d[1], { size: 19 })], { width: w[1] }),
    cell([cp(d[2], { size: 19 })], { width: w[2] }),
    cell([cp(d[3], { size: 19 })], { width: w[3] }),
  ]})));
  children.push(makeTable(w, rows));
}
children.push(body("总结语：借民歌写深情，借物象写胸怀，借圣君写承诺——全诗句句是情，句句是政。私人的慨叹经此三「借」上升为政治家的襟抱，既含蓄典雅又刚健慷慨，这就是「建安风骨」。", { bold: true, color: MAIN }));

children.push(H3("2. 情感总结：三忧一志——以「忧」写「志」，「我忧」实为「我要」"));
{
  const w = [1400, 3600, 4746];
  const rows = [headerRow(["情感", "对应章句", "内涵辨析（忧的背后是什么）"], w)];
  const data = [
    ["一忧：人生短暂", "第 1—2 章", "不是畏死贪生、及时行乐，而是忧来日无多、功业未竟——忧得越重，越见其「要抓紧干」的紧迫"],
    ["二忧：贤才未至", "第 3—5 章", "不是个人离愁，而是忧贤才散落四方、尚未归我——思念有多「悠悠」，求贤之心就有多切"],
    ["三忧：功业未就", "第 6—7 章", "乌鹊择枝，忧的是天下贤士无所依托，更是一统大业尚待其人——忧人即是忧世"],
    ["一志：天下归心", "第 8 章", "山海之量、周公之诚，终章豁然开朗：所有「忧」都收束到「志」——广纳英才、平定天下"],
  ];
  data.forEach(d => rows.push(new TableRow({ cantSplit: true, children: [
    cell([cp(d[0], { bold: true, color: d[0].startsWith("一志") ? RED : MAIN, align: AlignmentType.CENTER, size: 19 })], { width: w[0], fill: LIGHT_BG }),
    cell([cp(d[1], { size: 19 })], { width: w[1] }),
    cell([cp(d[2], { size: 19 })], { width: w[2] }),
  ]})));
  children.push(makeTable(w, rows));
}
children.push(P([R("情感曲线：叹（人生）→ 思（贤才）→ 迎（嘉宾）→ 忧（难掇）→ 聚（旧恩）→ 唤（乌鹊）→ 誓（归心）。",
  { bold: true, color: RED, size: 22 })], { align: AlignmentType.CENTER, before: 80, after: 40 }));
children.push(body("「忧」是表象与动力，「志」是内核与归宿：因志大而忧深，忧愈深而志愈坚。读此诗当读出「我要」——要时间、要人才、要天下太平；而非「我颓」——借酒避世、及时行乐。"));

// 八、板书
children.push(H1("八、板书设计"));
{
  const w = [9746];
  const lines = [
    ["短 歌 行    曹 操", true, MAIN],
    ["（汉乐府旧题·四言古诗·建安风骨）   前 4 章：师精讲授法 ｜ 后 4 章：生探究自悟", false, "555555"],
    ["", false, DARK],
    ["鉴赏四步法：咬文嚼字 → 辨识手法 → 还原情境（时·地·人·事·态）→ 体悟情志", true, MAIN],
    ["", false, DARK],
    ["三「借」：借民歌（子衿·鹿鸣）  借物象（朝露·明月·乌鹊）  借圣君（山海·周公）", false, DARK],
    ["", false, DARK],
    ["三忧：忧人生  忧贤才  忧功业  ——→  一志：天下归心", true, RED],
    ["忧而不颓 · 悲中见壮：「我忧」实为「我要」", false, DARK],
  ];
  const rows = [new TableRow({ cantSplit: true, children: [cell(
    lines.map(([t, b, c]) => new Paragraph({
      alignment: AlignmentType.CENTER, spacing: { line: 360, after: 40 },
      children: [new TextRun({ text: t, bold: b, size: b ? 24 : 21, color: c })],
    })), { width: w[0], fill: LIGHT_BG })] })];
  children.push(makeTable(w, rows));
}

// 九、听证三星评价
children.push(H1("九、当堂听证「三星」评价标准（投影 · 举牌/手势）"));
{
  const w = [2400, 5326, 2020];
  const rows = [headerRow(["星项", "达标要求", "判定方式"], w)];
  const data = [
    ["★ 证据星", "三条证据全部来自诗句原文，能点出关键词（如「悠悠」「枉」「何枝」）", "证据不出自诗句则不亮星"],
    ["★ 手法星", "比兴 / 用典 / 引用判断正确，能说出表达效果", "手法说错则不亮星"],
    ["★ 说服星", "结论有「忧中见志」的依据，陈词清楚、能回应质询", "脱离文本贴标签则不亮星"],
  ];
  data.forEach(d => rows.push(new TableRow({ cantSplit: true, children: [
    cell([cp(d[0], { bold: true, color: MAIN, align: AlignmentType.CENTER, size: 20 })], { width: w[0], fill: LIGHT_BG }),
    cell([cp(d[1], { size: 20 })], { width: w[1] }),
    cell([cp(d[2], { size: 19, align: AlignmentType.CENTER })], { width: w[2] }),
  ]})));
  children.push(makeTable(w, rows));
}

// 十、作业
children.push(H1("十、分层作业"));
children.push(bullet([R("基础类（必做）：", { bold: true }), R("背诵默写全诗；在笔记本完成手法总结表（三「借」7 处）。")]));
children.push(bullet([R("提升类（必做）：", { bold: true }), R("完成文学短评一篇（300 字左右，题目与支架见附录三），对接本单元「学写文学短评」任务。")]));
children.push(bullet([R("拓展类（选做）：", { bold: true }), R("① 用附录四提示词生成「月明星稀，乌鹊南飞」意境图或 30 秒数字人朗诵视频，附 50 字说明；② 比较《三国演义》中的曹操与诗中曹操，写 200 字札记。")]));

// 十一、反思
children.push(H1("十一、教学反思（课后填写）"));
children.push(bullet("对比诵读是否有效纠正了「读颓」的倾向？学生尾章能否自然读出「我要」的力量？"));
children.push(bullet("「还原情境五要素」是否让学生真正写出了画面？AI 生图是验证了情境还是替代了思考？"));
children.push(bullet("后 4 章探究中，学生能否独立完成「手法→政治理想」的分析链？点拨是否越界？"));
children.push(bullet("当堂听证无任何打印材料，学生的主张与证据质量如何？文学短评中短评写法（叙议结合）落实没有？"));

// ================= 附录 =================
children.push(new Paragraph({ children: [new PageBreak()] }));

// 附录一：学习单（投影誊抄版）
children.push(H1("附录一  八章鉴赏学习单（投影使用 · 学生誊入笔记本）"));
children.push(body("方法：咬文嚼字 → 辨识手法 → 还原情境（何时·何地·何人·何事·何态）→ 体悟情志", { color: "555555" }));
{
  const w = [720, 2420, 2100, 2206, 2300];
  const rows = [headerRow(["章", "原句", "关键词", "表达技巧", "情志（一个字词概括）"], w)];
  const filled = [
    ["1", "对酒当歌……去日苦多", "几何、苦多、朝露", "比兴（朝露喻人生短促）", "忧——时不我待（示范）"],
    ["2", "慨当以慷……唯有杜康", "忧思难忘、杜康", "设问、借酒写忧", "忧思沉郁（示范）"],
    ["3", "青青子衿……沉吟至今", "悠悠、沉吟、君", "引用《子衿》翻新：恋歌→思贤", "求贤之思（示范）"],
    ["4", "呦呦鹿鸣……鼓瑟吹笙", "嘉宾、鼓瑟吹笙", "引用《鹿鸣》：以礼贤之主自任", "迎贤之诚（示范）"],
  ];
  filled.forEach(d => rows.push(new TableRow({ cantSplit: true, children: [
    cell([cp(d[0], { align: AlignmentType.CENTER, bold: true, color: MAIN })], { width: w[0], fill: "EAF1FB" }),
    cell([cp(d[1], { size: 18 })], { width: w[1] }),
    cell([cp(d[2], { size: 18 })], { width: w[2] }),
    cell([cp(d[3], { size: 18 })], { width: w[3] }),
    cell([cp(d[4], { size: 18 })], { width: w[4] }),
  ]})));
  [["5", "明明如月……不可断绝"], ["6", "越陌度阡……心念旧恩"], ["7", "月明星稀……何枝可依"], ["8", "山不厌高……天下归心"]].forEach(d =>
    rows.push(new TableRow({ cantSplit: true, children: [
      cell([cp(d[0], { align: AlignmentType.CENTER, bold: true, color: MAIN })], { width: w[0] }),
      cell([cp(d[1], { size: 18 })], { width: w[1] }),
      cell([cp(" ")], { width: w[2] }), cell([cp(" ")], { width: w[3] }), cell([cp(" ")], { width: w[4] }),
    ]})));
  children.push(makeTable(w, rows));
}

// 附录二：当堂听证指令
children.push(H1("附录二  当堂「诗证·正名」听证指令（投影版 · 无需打印）"));
children.push(H3("辩题"));
children.push(body("为曹操摘掉「宁教我负天下人，休教天下人负我」的标签——诗中的曹操，究竟是怎样的人？"));
children.push(H3("第 1 步：3 分钟当堂起草（写在练习本上）"));
children.push(new Paragraph({
  spacing: { line: 340, after: 40 }, indent: { left: 300 },
  shading: { fill: WARN_BG, type: ShadingType.CLEAR, color: "auto" },
  border: { left: { style: BorderStyle.SINGLE, size: 18, color: "C55A11", space: 8 } },
  children: [new TextRun({
    text: "我的主张：诗中的曹操是一位 ________________ 的人。",
    size: 21 })],
}));
children.push(new Paragraph({
  spacing: { line: 340, after: 40 }, indent: { left: 300 },
  shading: { fill: WARN_BG, type: ShadingType.CLEAR, color: "auto" },
  border: { left: { style: BorderStyle.SINGLE, size: 18, color: "C55A11", space: 8 } },
  children: [new TextRun({
    text: "证据①：「____________」中的「____」（手法：____），可见 ____________。",
    size: 21 })],
}));
children.push(new Paragraph({
  spacing: { line: 340, after: 40 }, indent: { left: 300 },
  shading: { fill: WARN_BG, type: ShadingType.CLEAR, color: "auto" },
  border: { left: { style: BorderStyle.SINGLE, size: 18, color: "C55A11", space: 8 } },
  children: [new TextRun({
    text: "证据②：「____________」中的「____」（手法：____），可见 ____________。",
    size: 21 })],
}));
children.push(new Paragraph({
  spacing: { line: 340, after: 60 }, indent: { left: 300 },
  shading: { fill: WARN_BG, type: ShadingType.CLEAR, color: "auto" },
  border: { left: { style: BorderStyle.SINGLE, size: 18, color: "C55A11", space: 8 } },
  children: [new TextRun({
    text: "证据③：「____________」中的「____」（手法：____），可见 ____________。",
    size: 21 })],
}));
children.push(H3("第 2 步：2 分钟组内互证"));
children.push(body("两人一组轮流即席陈述；听者只做一件事——追问：「这个结论出自哪一个词、哪一个典故？」答不上来的证据划掉重找。"));
children.push(H3("第 3 步：每组 1 人即席陈词 1 分钟 + 接受 30 秒质询"));
children.push(bullet("开场句式：「我方认为……第一，请看诗句……；第二，从……一词可见……；第三，……之典表明……。综上，他不是……而是……。」"));
children.push(bullet("可用素材仅限：课本、课堂笔记（不许翻阅其他资料）；三星标准即第九节。"));

// 附录三：求贤令+文学短评
children.push(H1("附录三  《求贤令》节选 与 文学短评写作支架（投影 / 课后写作）"));
children.push(H3("1. 拓展资料：曹操《求贤令》（建安十五年）节选"));
children.push(new Paragraph({
  spacing: { line: 340, after: 60 }, indent: { left: 300 },
  children: [new TextRun({ text: "自古受命及中兴之君，曷尝不得贤人君子与之共治天下者乎！……今天下得无有被褐怀玉而钓于渭滨者乎？……二三子其佐我明扬仄陋（zé lòu，举荐出身卑微的人才），唯才是举，吾得而用之。", size: 21, italics: true })],
}));
children.push(body("对读印证：「鼓瑟吹笙」之诺——正是「与之共治天下」；「山不厌高，海不厌深」——正是「明扬仄陋、唯才是举」；「周公吐哺」——正是求贤的至诚姿态。诗是深情版的求贤令，令是政令版的短歌行。", { color: GREEN, bold: false }));

children.push(H3("2. 文学短评任务（对接本单元「学写文学短评」）"));
children.push(bullet([R("题目（四选一或自拟）：", { bold: true })]));
children.push(bullet("① 一个「忧」字见英雄——《短歌行》情感辨", 1));
children.push(bullet("② 从《子衿》到《鹿鸣》：曹操怎样「借」诗求贤", 1));
children.push(bullet("③ 诗与令之间——读《短歌行》与《求贤令》", 1));
children.push(bullet("④ 乌鹊择枝与周公吐哺：曹操的一场人才告白", 1));
children.push(bullet([R("写作支架（三字诀）：", { bold: true, color: RED })]));
children.push(bullet([R("小——", { bold: true }), R("切口要小：只评一个词、一种手法或一对矛盾（如只评「枉」字、只评「乌鹊」之喻），不做全诗翻译。")], 1));
children.push(bullet([R("证——", { bold: true }), R("以诗为证：至少 3 处引用原句，引后必析，说出「写了什么—怎么写—好在哪」。")], 1));
children.push(bullet([R("议——", { bold: true }), R("叙议结合：描情境（叙）是手段，评情志、评手法（议）是目的；段段有观点句，结尾回应「这是怎样一个曹操」。")], 1));
children.push(body("评价要点：观点明确（20）＋ 引诗准确、分析扣住「手法如何传达理想」（40）＋ 叙议结合、结构清晰（25）＋ 有自己的发现（15），300 字左右。", { color: "555555" }));

// 附录四：AI 提示词包
children.push(H1("附录四  AI 情境生成提示词包（课后选做 · 先有情境句，再用 AI）"));
children.push(H3("情境公式"));
children.push(P([R("时代背景 + 地点 + 时间天气 + 人物动作神态 + 氛围基调 = 可生成的「情境表述」", { bold: true, color: MAIN, size: 22 })], { align: AlignmentType.CENTER, after: 80 }));
children.push(H3("范例 1 · 文生图（第 1 章 / 第 7 章）"));
children.push(bullet("「中国水墨画与电影光影结合风格，东汉末年江畔军营秋夜，一轮明月、疏星、江面战船剪影，大帐外夜宴灯火，一位年逾五旬、身形苍劲的将军停杯仰头望月，案上酒樽与竹简，草叶上有朝露；氛围苍凉沉郁而壮阔，冷暖光对比，横版 16:9，留白」"));
children.push(H3("范例 2 · 数字人短视频"));
children.push(bullet("「生成一位东汉末年政治家形象的数字人：须发斑白、气度沉雄，着汉代深衣，于月夜军营宴上举杯而立；用苍劲低沉、先抑后扬的声音朗诵《短歌行》第 1、8 章；前四句语速缓慢、神色苍凉，末两句渐强、目光坚定，背景鼓乐由弱转强，时长 30 秒。」"));
children.push(H3("范例 3 · AI 配音情绪脚本（直接标给配音工具）"));
children.push(bullet("「对酒当歌（沉、缓），人生几何（放轻、延长）……青青子衿（柔、自语），悠悠我心（绵长、渐强）……山不厌高（渐强、稳重），海不厌深（再强），周公吐哺（庄重），天下归心（饱满、上扬、余韵坚定）。」"));
children.push(bullet([R("使用三原则：", { bold: true, color: RED }), R("① 先读诗、先写情境句，再用 AI；② AI 的画面 / 典故必须回到原诗验证（指出 AI 一处「不符合诗意」的地方可加分）；③ AI 是助手，作品由本人署名负责。")]));

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
