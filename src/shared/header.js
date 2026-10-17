/** 加载项插入签名后打在邮件上的头。Exchange 兜底规则看到它就不再追加团队信息（见 fallback.js） */
export const SIGNATURE_HEADER = "X-ACS-Signature";
