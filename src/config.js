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
    termsText: "Confidentiality Statement & Terms and Conditions",
    // 免责声明（一行）：disclaimerPrefix + 链接文字(termsText)
    disclaimerPrefix: "This email is subject to our",
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
    // 不需要底部徽章 / 圆环就设为 null
  },

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
  },
  fontFamily: "Arial, Helvetica, sans-serif",
};
