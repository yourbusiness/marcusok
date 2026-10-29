/**
 * MS-OFFCRYPTO Agile 加密 pptx 生成器（测试 fixture 用）。
 *
 * 真实二进制加密样本（非 mock）：`cfb` 写 CFB 容器 + Node `crypto` 实现
 * [MS-OFFCRYPTO] 2.3.6.2 密钥派生与 AES-CBC。已闭环验证上游可解密（正确
 * 密码解出原文字节、错误密码报 invalid-password，见 docs/ppt-preview-
 * spike-record.md S2）。
 *
 * 两个非显然的上游兼容细节（spec 未明说、实测定案）：
 * 1. verifier 哈希作用于**含 PKCS7 padding 的明文**——上游解密不去
 *    padding，比较的是 padded 明文的 SHA-512；
 * 2. EncryptedPackage 中间段（4096B 对齐）必须 no-padding：PKCS7 会把
 *    中间段变 4112B 破坏段边界；仅末段 PKCS7，上游按 totalSize 截断明文。
 */
import {
  createCipheriv,
  createHash,
  createHmac,
  randomBytes,
} from "node:crypto";
import * as CFB from "cfb";

// blockKey 常量（[MS-OFFCRYPTO] 2.3.6.2/2.3.6.5；与 msoffcrypto-tool 参考实现核对）
const BLK = {
  verifierHashInput: Buffer.from("fea7d2763b4b9e79", "hex"),
  verifierHashValue: Buffer.from("d7aa0f6d3061344e", "hex"),
  keyValue: Buffer.from("146e0be7abacd0d6", "hex"),
  hmacKey: Buffer.from("5fb2ad010cb9e1f6", "hex"),
  hmacValue: Buffer.from("a0677f02b22c8433", "hex"),
} as const;

const SHA512 = (b: Buffer) => createHash("sha512").update(b).digest();

/**
 * 密钥派生：H0=SHA512(salt+pwd[UTF-16LE])；Hn=SHA512(LE32(i)+Hn-1)（i 从
 * 0 起，最后一轮 = spinCount-1）；key=SHA512(Hn+blockKey) 截取 keyBits/8
 * （SHA-512 摘要 64B ≥ 32B，无 0x36 填充分支）。
 */
function deriveKey(
  password: string,
  salt: Buffer,
  spinCount: number,
  blockKey: Buffer,
  keyBits: number,
): Buffer {
  let h = SHA512(Buffer.concat([salt, Buffer.from(password, "utf16le")]));
  const iter = Buffer.alloc(4);
  for (let i = 0; i < spinCount; i++) {
    iter.writeUInt32LE(i, 0);
    h = SHA512(Buffer.concat([iter, h]));
  }
  return SHA512(Buffer.concat([h, blockKey])).subarray(0, keyBits / 8);
}

/** AES-256-CBC 加密（PKCS7；noPad 用于整段对齐的中间段）。 */
function aesCbc(key: Buffer, iv: Buffer, plain: Buffer, noPad = false): Buffer {
  const c = createCipheriv("aes-256-cbc", key, iv);
  if (noPad) c.setAutoPadding(false);
  return Buffer.concat([c.update(plain), c.final()]);
}

export interface AgileOptions {
  /** KDF 迭代次数（真实 Office 默认 100000；测试默认 1000 提速，spec 允许任意值）。 */
  spinCount?: number;
}

/** 把明文 pptx 字节封进 Agile 加密的 OLE2 容器（ EncryptionInfo + EncryptedPackage）。 */
export function encryptAgile(
  plainBytes: Uint8Array,
  password: string,
  opts: AgileOptions = {},
): Uint8Array {
  const spinCount = opts.spinCount ?? 1000;
  const plain = Buffer.from(plainBytes);
  const keySalt = randomBytes(16); // p:encryptedKey 的 saltValue
  const dataSalt = randomBytes(16); // keyData 的 saltValue
  const secretKey = randomBytes(32); // 内容加密密钥（存入 encryptedKeyValue）
  const verifierInput = randomBytes(16);
  const hmacKey = randomBytes(64);

  // password 派生 key：verifier/secret 用 encryptedKey salt，HMAC 用 keyData salt
  const kVerifierIn = deriveKey(
    password,
    keySalt,
    spinCount,
    BLK.verifierHashInput,
    256,
  );
  const kVerifierHash = deriveKey(
    password,
    keySalt,
    spinCount,
    BLK.verifierHashValue,
    256,
  );
  const kSecret = deriveKey(password, keySalt, spinCount, BLK.keyValue, 256);
  const kHmacKey = deriveKey(password, dataSalt, spinCount, BLK.hmacKey, 256);
  const kHmacVal = deriveKey(password, dataSalt, spinCount, BLK.hmacValue, 256);

  // verifier：哈希输入为 padded 明文（上游解密不去 padding）
  const pad16 = (b: Buffer): Buffer => {
    const pad = 16 - (b.length % 16);
    return Buffer.concat([b, Buffer.alloc(pad, pad)]);
  };
  const encVerifierInput = aesCbc(kVerifierIn, keySalt, verifierInput);
  const encVerifierHash = aesCbc(
    kVerifierHash,
    keySalt,
    SHA512(pad16(verifierInput)),
  );
  const encSecret = aesCbc(kSecret, keySalt, secretKey);

  // 内容：LE64(明文长度) + 分段 CBC（4096B/段，IV_i=SHA512(dataSalt+LE32(i))[:16]）
  const SEG = 4096;
  const segCount = Math.ceil(plain.length / SEG) || 0;
  const encBody: Buffer[] = [];
  const idx = Buffer.alloc(4);
  for (let i = 0; i < segCount; i++) {
    idx.writeUInt32LE(i, 0);
    const iv = SHA512(Buffer.concat([dataSalt, idx])).subarray(0, 16);
    const chunk = plain.subarray(
      i * SEG,
      Math.min((i + 1) * SEG, plain.length),
    );
    encBody.push(aesCbc(secretKey, iv, chunk, chunk.length === SEG));
  }
  const sizeHeader = Buffer.alloc(8);
  sizeHeader.writeBigUInt64LE(BigInt(plain.length), 0);
  const encryptedPackage = Buffer.concat([sizeHeader, ...encBody]);

  // dataIntegrity：上游解密不校验 HMAC（spike 核实），按 spec 生成——
  // PowerPoint 等严格实现可校验
  const encHmacKey = aesCbc(kHmacKey, dataSalt, hmacKey);
  const hmacValue = createHmac("sha512", hmacKey)
    .update(encryptedPackage)
    .digest();
  const encHmacValue = aesCbc(kHmacVal, dataSalt, hmacValue);

  const b64 = (b: Buffer) => b.toString("base64");
  const xml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" ` +
    `xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password" ` +
    `xmlns:c="http://schemas.microsoft.com/office/2006/keyEncryptor/certificate">` +
    `<keyData saltSize="16" blockSize="16" keyBits="256" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="${b64(dataSalt)}"/>` +
    `<dataIntegrity encryptedHmacKey="${b64(encHmacKey)}" encryptedHmacValue="${b64(encHmacValue)}"/>` +
    `<keyEncryptors><keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">` +
    `<p:encryptedKey spinCount="${spinCount}" saltSize="16" blockSize="16" keyBits="256" hashSize="64" ` +
    `cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="${b64(keySalt)}" ` +
    `encryptedVerifierHashInput="${b64(encVerifierInput)}" encryptedVerifierHashValue="${b64(encVerifierHash)}" ` +
    `encryptedKeyValue="${b64(encSecret)}"/>` +
    `</keyEncryptor></keyEncryptors></encryption>`;

  // EncryptionInfo 流：LE16(4)+LE16(4)+LE32(0)+XML（Agile v4 头）
  const ver = Buffer.alloc(12);
  ver.writeUInt16LE(4, 0);
  ver.writeUInt16LE(4, 2);
  const encryptionInfo = Buffer.concat([ver, Buffer.from(xml, "utf8")]);

  const cfb = CFB.utils.cfb_new();
  CFB.utils.cfb_add(cfb, "/EncryptionInfo", encryptionInfo);
  CFB.utils.cfb_add(cfb, "/EncryptedPackage", encryptedPackage);
  return new Uint8Array(CFB.write(cfb));
}
