/**
 * ACS 签名配置 —— 公司统一信息都在这里改，改完重新 build + 推送即可全员生效。
 * 个人信息（姓名、手机、电话、分机、职位）从 Entra ID 自动读取，不在这里写。
 */
export const SIGNATURE_CONFIG = {
  company: {
    name: "Aion Cargo Solutions",
    shortName: "ACS",
    website: "www.aioncargo.com",
    websiteUrl: "https://www.aioncargo.com",
    // Entra 里没有填办公电话时使用的总机号码
    mainPhone: "+61 2 9160 2300",
    address: "Suite 4.06, 55 Miller St, Pyrmont NSW 2009",
    // 点击地址时打开的地图链接（留空则地址不加链接）
    addressUrl: "https://maps.google.com/?q=Suite+4.06,+55+Miller+St,+Pyrmont+NSW+2009",
    // 保密声明 & 条款链接 —— 请替换为你们真实的地址
    termsUrl: "https://www.aioncargo.co.nz/au-tc/",
    termsText: "Terms and Conditions",
    // 免责声明（一行）：disclaimerPrefix + 链接文字(termsText)
    disclaimerPrefix: "The content of this email, including any ACS AU quotation, offer, booking, and other services, is subject to ACS AU",
  },

  // 签名上方的结束语；设为 "" 则不插入
  signOff: "Kind regards,",
  // 精简签名是否也带结束语
  signOffInShort: true,

  // 完整签名是否在姓名下方显示职位
  showTitleInFull: true,
  // 精简签名是否显示职位
  showTitleInShort: true,
  // 精简签名是否在公司名称后显示澳洲公司地址
  showAddressInShort: true,

  /**
   * 共享邮箱（团队邮箱）：发件人 (From) 是下面某个地址时，自动改用团队版签名 ——
   * 发件人姓名 + 团队名 + 团队电话 / 分机 + 团队邮箱，默认不显示个人手机。
   *  - addresses ：这个邮箱的全部地址（不区分大小写）。第一个显示在签名里，其余只用来识别发件人；
   *                Exchange 里的主地址和别名都要写上，否则用另一个地址发信时识别不到
   *  - team      ：签名里显示的团队名
   *  - phone/ext ：团队电话和分机
   *  - showMobile：设为 true 时，这个团队的签名也显示发件人本人的手机
   * 不在这张表里的共享邮箱，签名仍然是发件人本人的。新增 / 修改团队：改这里 → 提交即可，不需要重新上传 manifest。
   */
  sharedMailboxes: [
    { addresses: ["accounts@aioncargo.com.au", "au.accounts@aioncargo.com"], team: "ACS AU Accounts", phone: "+61 2 9160 2300", ext: "820" },
    { addresses: ["admin@aioncargo.com.au", "au.admin@aioncargo.com"], team: "ACS AU Admin", phone: "+61 2 9160 2300", ext: "602" },
    { addresses: ["customercare@aioncargo.com.au", "au.customercare@aioncargo.com"], team: "ACS AU Customer Care", phone: "+61 2 9160 2300", ext: "601" },
    { addresses: ["exportair@aioncargo.com.au", "au.exportair@aioncargo.com"], team: "ACS AU Export Air", phone: "+61 2 9160 2300", ext: "822" },
    { addresses: ["exportsea@aioncargo.com.au", "au.exportsea@aioncargo.com"], team: "ACS AU Export Sea", phone: "+61 2 9160 2300", ext: "822" },
    { addresses: ["importair@aioncargo.com.au", "au.importair@aioncargo.com"], team: "ACS AU Import Air", phone: "+61 2 9160 2300", ext: "821" },
    { addresses: ["importsea@aioncargo.com.au", "au.importsea@aioncargo.com"], team: "ACS AU Import Sea", phone: "+61 2 9160 2300", ext: "821" },
  ],

  /**
   * 其他办公室：Entra "办公室"(officeLocation) 包含 match 里任一文字（不区分大小写）的用户，用这里的设置覆盖上面的默认（澳洲）设置。
   * 只写和澳洲不同的项；company / labels / images 按字段合并，其余整项替换。Logo、版式、颜色沿用默认。
   * 匹配不到（或 Entra 没填办公室）的用户一律用默认设置。新增 / 修改办公室：改这里 → 提交即可，不需要重新上传 manifest。
   */
  offices: [
    {
      match: ["shanghai", "上海"],
      company: {
        name: "艾朗国际物流（上海）有限公司",
        mainPhone: "+86 21 65050600",
        address: "上海市虹口区衡水路151号前海人寿大厦813室",
        // 国内打不开 Google 地图，地址不加链接
        addressUrl: "",
        termsUrl: "https://www.aioncargo.co.nz/wp-content/uploads/2023/11/ACS-Terms-and-Conditions.pdf",
        termsText: "Terms and Conditions",
        disclaimerPrefix: "The content of this email, including any ACS quotation, offer, booking, and other services, is subject to ACS",
      },
      // 完整签名在职位下方显示公司中文名
      showCompanyInFull: true,
      labels: { web: "Website" },
      // 暂不放 WCA 徽章
      images: { badge: null, banner: null },
      // Arial 没有中文字形，中文回退到雅黑 / 苹方
      fontFamily: "Arial, Helvetica, 'Microsoft YaHei', 'PingFang SC', sans-serif",
    },
    {
      match: ["new zealand", "auckland"],
      company: {
        mainPhone: "+64 9 254 4900",
        address: "40 Airpark Drive, Mangere, Auckland 2022",
        addressUrl: "https://maps.google.com/?q=40+Airpark+Drive,+Mangere,+Auckland+2022",
        termsUrl: "https://www.aioncargo.co.nz/wp-content/uploads/2023/11/ACS-Terms-and-Conditions.pdf",
        termsText: "Terms and Conditions",
        disclaimerPrefix: "The content of this email, including any ACS quotation, offer, booking, and other services, is subject to ACS",
      },
      // 新西兰习惯把直线电话写作 DDI
      labels: { phone: "DDI" },
      // 澳洲的 WCA 徽章印着澳洲会员号 152015，新西兰用自己的三个徽章（放在免责声明下面一行）。
      // shortcut: 图片是从 Ivan 旧签名截图里裁出来的（白底，约 2 倍分辨率），拿到原图后替换同名文件即可
      images: {
        banner: null,
        iata: { file: "iata-badge.png", width: 55, height: 39, alt: "IATA" },
        cbaff: { file: "cbaff-badge.png", width: 56, height: 56, alt: "CBAFF New Zealand" },
        wcaNz: { file: "wca-nz-badge.png", width: 77, height: 57, alt: "WCA Advanced Professionals" },
      },
      footerBadges: ["iata", "cbaff", "wcaNz"],
    },
  ],

  /**
   * 个人附加说明：按本人邮箱（不区分大小写），以红字显示在结束语和姓名之间，完整签名和精简签名都有。
   * 每项一行；写成 { text: "...", bold: true } 则加粗。用共享邮箱发信时不显示（那是团队签名）。
   * 新增 / 修改：改这里 → 提交即可，不需要重新上传 manifest。
   */
  personalNotes: {
    "candy.wang@aioncargo.com": [
      "所有提单信息都请提供一次性补料，任何更改都会收取改单费RMB500",
      { text: "马士基VGM仅限一次性申报，更新会有罚金USD300", bold: true },
    ],
    "caroline@aioncargo.com": [
      "所有提单信息都请提供一次性补料，任何更改都会收取改单费RMB500",
      "马士基VGM仅限一次性申报，更新会有罚金USD300",
      { text: "为了贵司货物顺利出运，请务必核对放舱信息(包括 卸货港 目的港 箱型 船期 码头等)如有问题 请及时联系更改 如无异议视作默认放舱内容 感谢您的配合!", bold: true },
    ],
    "phoebe.wang@aioncargo.com": [
      "需要开票请提供开票抬头，项目名称默认开国际货物运输代理服务费，如对发票有要求请在开发票前提出；付款后麻烦提供水单，我司将在复核后销账，谢谢配合~",
    ],
  },

  // 分机号来源：先从 Entra "办公电话" 里解析（如 "+61 2 9160 2300 EXT 601" / "x601"），
  // 解析不到再读这个扩展属性（Exchange 自定义属性 1 = extensionAttribute1）
  extensionAttribute: "extensionAttribute1",

  /**
   * 图片模式：
   *  - "embed"：图片以内嵌附件（CID）方式插入，任何客户端都能直接显示，不会被“阻止下载图片”拦截（推荐）
   *  - "link" ：图片引用 baseUrl/assets/ 上的外链，邮件更小，但收件人 Outlook 可能默认不显示图片
   */
  imageMode: "embed",

  // 联系方式标签。注意："Email" 这一行也用于识别引用邮件里是否已有本人签名（profile.js bodyHasOwnSignature）
  labels: { mobile: "Mobile", phone: "Tel", email: "Email", web: "Web", address: "Address" },

  images: {
    // 文件放在 src/assets/ 下；width/height 是显示尺寸（源图为 2 倍分辨率，透明背景 + 白色描边以适配暗色模式）
    // Logo 高度与右侧姓名 + 联系方式信息块一致
    logo: { file: "acs-logo.png", width: 150, height: 91, alt: "ACS - Aion Cargo Solutions" },
    badge: { file: "wca-badge.png", width: 78, height: 52, alt: "WCA Advanced Professionals" },
    // ACS 品牌圆环标志：完整圆环（footer / name 位置用）
    mark: { file: "acs-mark.png", width: 40, height: 38, alt: "" },
    // 淡化的 1/4 圆环（带箭头那一段），放在联系方式右下角做背景装饰
    markCorner: { file: "acs-mark-corner.png", width: 48, height: 45, alt: "" },
    // 节日横幅：完整签名和精简签名最下面都显示，until 之后自动不再显示（按写邮件那台电脑 / 手机的时间）。
    // 只给澳洲：其他办公室在各自的 images 里设成 null。过期后可以删掉这一项和图片文件。
    banner: {
      file: "christmas-2026.png", width: 480, height: 88, until: "2027-01-02T00:00:00+11:00",
      alt: "Merry Christmas & Happy New Year. Office closed: 25, 26 & 28 Dec 2026 and 1 Jan 2027. Thank you for your support, from all of us at ACS AU.",
    },
    // 不需要底部徽章 / 圆环 / 横幅就设为 null
  },
  // 完整签名底部显示的徽章（images 里的名字，按顺序）。只有一个时放在免责声明左边；多个时另起一行放在免责声明下面
  footerBadges: ["badge"],

  /**
   * 品牌圆环标志的位置（完整签名）：
   *  - "overlay"：淡化的 1/4 圆环压在地址那行文字后面（推荐）。
   *               新版 Outlook / 网页版 Outlook / Gmail / Apple Mail 正常显示；
   *               经典 Outlook（Word 引擎）不支持定位，会自动隐藏圆环，其余部分不受影响。
   *  - "corner" ：同样的圆环，但放在联系方式右侧单独一格、不与文字重叠（所有客户端都显示）
   *  - "footer" ：底部分隔线右端的完整圆环
   *  - "name"   ：姓名这一行最右侧的完整圆环
   *  - "none"   ：不显示
   * 精简签名不放图片（每封回复都会多一个内嵌附件），所以圆环只出现在完整签名里。
   */
  markPosition: "overlay",
  // overlay 位置微调：距联系方式区域右边 / 下边多少像素（right:0 = 圆环右边缘与地址末尾 "2009" 对齐）
  markOverlayOffset: { right: 0, bottom: -2 },

  // 精简签名里也放这个淡化圆环（网址那一行附近）。代价：每封回复多一个 ~4KB 的内嵌图片
  markInShort: true,
  markShortOffset: { right: 0, bottom: 16 },

  // 字号（px）—— 想整体再紧凑或再大一点，改这里就行
  fontSizes: { name: 17, title: 11, row: 11, label: 10, rowLine: 16, footer: 9.5 },

  /**
   * 回复 / 转发规则：
   * - 新邮件：完整签名
   * - 回复：如果本人在这个邮件会话里还没发过邮件 → 完整签名；已经发过 → 精简签名
   * - 转发：forwardAlwaysFull=true 时始终用完整签名（转给新的人通常需要完整联系方式）
   * - 共享邮箱回复：引用的历史邮件里已经有这个团队的签名 → 精简签名，否则完整签名（不看是谁发的）
   */
  forwardAlwaysFull: true,

  colors: {
    text: "#1a1a1a",
    value: "#333333", // 联系方式内容
    muted: "#6b6b6b",
    label: "#D70F44", // 联系方式前的标签（Mobile / Tel / Email / Web / Address）
    link: "#1a1a1a",
    footer: "#8c8c8c", // 底部免责声明
    divider: "#D70F44", // Logo 与信息之间的竖线
    hairline: "#e2e2e2", // 底部分隔细线
    accent: "#D70F44", // ACS Logo 红色
    note: "#D70F44", // 个人附加说明（personalNotes）
  },
  fontFamily: "Arial, Helvetica, sans-serif",
};
