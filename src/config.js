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
    logo: { file: "acs-logo.png", width: 248, height: 150, alt: "ACS - Aion Cargo Solutions" },
    badge: { file: "wca-badge.png", width: 90, height: 60, alt: "WCA Advanced Professionals" },
    // 不需要底部徽章就设为 null
  },

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
