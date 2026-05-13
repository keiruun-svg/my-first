/**
 * OJC 자동 품번 생성 로직  (AJW 통신선로개발팀)
 * 품번 체계: A14-[①OJC종류]-[②코어종류③심선수④재질]-[⑤커넥터A⑥커넥터B]-[⑦길이코드⑧마킹]
 *
 * 작성일: 2026-05-13
 * VS Code 내부 서버화 예정
 */

// ────────────────────────────────────────────────────────────────────────────
// 타입 정의
// ────────────────────────────────────────────────────────────────────────────

export interface PartFields {
  ojcType:    string | null;  // ① 위치 4
  coreType:   string | null;  // ② 위치 5
  coreCount:  string | null;  // ③ 위치 6
  material:   string | null;  // ④ 위치 7
  connectorA: string | null;  // ⑤ 위치 8
  connectorB: string | null;  // ⑥ 위치 9
  lengthCode: string | null;  // ⑦ 위치 10~12
  marking:    string | null;  // ⑧ 위치 13
}

export interface GenerateResult {
  partNumber: string | null;
  fields:     PartFields;
  missing:    string[];
  error:      string | null;
}

// ────────────────────────────────────────────────────────────────────────────
// ① OJC 종류 코드표  (위치 4)
// ────────────────────────────────────────────────────────────────────────────
export const OJC_TYPE: Record<string, string> = {
  A: "Simplex 2.0mm        — SOJC / OJC-A1 SP",
  B: "Simplex 3.0mm",
  C: "Duplex                — DOJC / OJC-A1 DP",
  D: "MOJC 0.9mm           (현재 미사용)",
  E: "MOJC 2.0mm / OJC-C2  — MOJC-SM / OJC-C2",
  F: "Pigtail 0.9mm",
  G: "Pigtail 2.0mm",
  H: "Pigtail 3.0mm",
  I: "Pigtail 성단용",
  J: "리본케이블",
  K: "Drop cable 3.0mm     — DROP-CABLE",
  L: "Drop cable 5.0mm     (현재 미사용)",
  M: "Optical Cable 3.0mm  — Optical Cable Parts",
  N: "FLAT 3mm×2mm",
};

// ────────────────────────────────────────────────────────────────────────────
// ② 코어 종류 코드표  (위치 5)
// ────────────────────────────────────────────────────────────────────────────
export const CORE_TYPE: Record<string, string> = {
  A: "SMF 652D",
  B: "SMF 657.A1",
  C: "SMF 657.A2",
  D: "SMF 657.B3  ← 품목명 SM 표기 시 기본값",
  E: "OM1",
  F: "OM2",
  G: "OM3",
  H: "OM4",
  I: "OM5",
};

// ────────────────────────────────────────────────────────────────────────────
// ③ 심선수 코드표  (위치 6)
//    ※ 코드값 ≠ 심수 — 직접 대응 아님
// ────────────────────────────────────────────────────────────────────────────
export const CORE_COUNT: Record<string, number> = {
  "1":  1,  "2":  2,  "3":  4,  "4":  6,
  "5":  7,  "6":  8,  "7": 10,  "8": 12,
  "9": 13,  "A": 14,  "B": 15,  "C": 16,
  "D": 17,  "E": 18,  "F": 20,  "G": 22,
  "H": 24,  "I": 36,  "J": 48,
};
// 심수 → 코드 역변환
export const CORE_COUNT_INV: Record<number, string> =
  Object.fromEntries(Object.entries(CORE_COUNT).map(([k, v]) => [v, k]));

// ────────────────────────────────────────────────────────────────────────────
// ④ 재질 코드표  (위치 7)
// ────────────────────────────────────────────────────────────────────────────
export const MATERIAL: Record<string, string> = {
  "0": "PVC YELLOW",  "1": "PVC BLUE",    "2": "PVC AQUA",
  "3": "PVC ORANGE",  "4": "PVC BLACK",   "5": "PVC WHITE",
  "6": "PVC 2색",     "7": "PVC 4색",     "8": "PVC 6색",
  "9": "PVC 8색",     "A": "PVC 12색",
  "B": "LSZH BLACK",  "C": "LSZH YELLOW",
  "D": "PU YELLOW",   "E": "PU BLUE",     "F": "PU AQUA",
  "G": "PU ORANGE",   "H": "PU BLACK",    "I": "PU WHITE",
  "J": "PU 2색",      "K": "PU 4색",      "L": "PU 6색",
  "M": "PU 8색",      "N": "PU 12색",
  "O": "PE YELLOW",   "P": "PE BLUE",     "Q": "PE AQUA",
  "R": "PE ORANGE",   "S": "PE BLACK",    "T": "PE WHITE",
  "U": "PE 2색",      "V": "PE 4색",      "W": "PE 6색",
  "X": "PE 8색",      "Y": "PE 12색",
};

// ────────────────────────────────────────────────────────────────────────────
// ⑤⑥ 커넥터 코드표  (위치 8, 9)
// ────────────────────────────────────────────────────────────────────────────
export const CONNECTOR: Record<string, string> = {
  "0": "SC/PC",        "1": "SC/APC",       "2": "SC/PC CLIP",
  "3": "SC/APC CLIP",  "4": "LC/PC",        "5": "LC/APC",
  "6": "LC/PC CLIP",   "7": "LC/APC CLIP",  "8": "FC/PC",
  "9": "FC/APC",       "A": "ST/PC",        "B": "ST/APC",
  "C": "MU/PC",        "D": "MU/APC",       "E": "MTRJ/PC",
  "F": "MPO A",        "G": "MPO B",        "H": "MPO C",
  "I": "MTP A",        "J": "MTP B",        "K": "MTP C",
};

// 커넥터 이름 → 코드
const CONNECTOR_ALIAS: Record<string, string> = {
  "SC/PC":               "0",  "SC/APC":              "1",
  "SC/PC CLIP":          "2",  "SC/PC,CLIP":          "2",
  "SC/APC CLIP":         "3",  "SC/APC,CLIP":         "3",
  "LC/PC":               "4",  "LC/APC":              "5",
  "LC/PC CLIP":          "6",  "LC/PC,CLIP":          "6",
  "LC DUPLEX KIT(CLIP)": "6",  "LC/APC CLIP":         "7",
  "LC/APC,CLIP":         "7",  "FC/PC":               "8",
  "FC/APC":              "9",  "ST/PC":               "A",
  "ST/APC":              "B",
};

// ────────────────────────────────────────────────────────────────────────────
// ⑧ 마킹 코드표  (위치 13)
// ────────────────────────────────────────────────────────────────────────────
export const MARKING: Record<string, string> = {
  N: "NON",  K: "KT",  L: "LG",
  O: "고려오트론",  P: "포앤티",  Z: "수출",  X: "케이블 사급",
};

// ────────────────────────────────────────────────────────────────────────────
// ⑦ 길이 코드 변환
// ────────────────────────────────────────────────────────────────────────────
const SPECIAL_LEN: Record<number, string> = {
  100: "100", 150: "150", 200: "200", 250: "250", 300: "300",
};

export function lengthToCode(m: number): string {
  if (SPECIAL_LEN[m]) return SPECIAL_LEN[m];
  const dec = (d: number) => String.fromCharCode("A".charCodeAt(0) + d - 1);
  if (m < 1) {
    const d = Math.round(m * 10);
    return `00${dec(d)}`;
  }
  if (m < 10) {
    const i = Math.floor(m), d = Math.round((m - i) * 10);
    return d === 0 ? `00${i}` : `0${i}${dec(d)}`;
  }
  if (m < 100) {
    const i = Math.floor(m), d = Math.round((m - i) * 10);
    return d === 0 ? `${String(i).padStart(2,"0")}J` : `${String(i).padStart(2,"0")}${dec(d)}`;
  }
  throw new Error(`지원하지 않는 길이: ${m}M`);
}

export function parseLength(spec: string): number | null {
  const s = spec.toUpperCase();
  const mm = s.match(/(\d+(?:\.\d+)?)\s*MM(?!\s*M)/);
  if (mm) return parseFloat(mm[1]) / 1000;
  const m = s.match(/(\d+(?:\.\d+)?)\s*M(?!M)/);
  if (m) return parseFloat(m[1]);
  return null;
}

// ────────────────────────────────────────────────────────────────────────────
// 자동 판별 함수들
// ────────────────────────────────────────────────────────────────────────────

/**
 * ⑧ 마킹 자동 판별 — 품목명 prefix 기반
 *
 *   OJC-A1- / OJC-C2-          → K (KT)
 *   SOJC- / DOJC- / MOJC-      → L (LG)
 *   DROP-CABLE / PIGTAIL- / 기타 → N (NON, 비계약 품목)
 */
export function detectMarking(productName: string): string {
  const n = productName.toUpperCase();
  if (n.startsWith("OJC-A1-") || n.startsWith("OJC-C2-")) return "K";
  if (n.startsWith("SOJC-") || n.startsWith("DOJC-") || n.startsWith("MOJC-")) return "L";
  return "N";
}

/**
 * ① OJC 종류 코드 자동 판별 — 품목명 기반
 *
 *   SOJC-      → A  (Simplex 2.0mm, LG)
 *   OJC-A1 SP  → A  (Simplex 2.0mm, KT)
 *   DOJC-      → C  (Duplex, LG)
 *   OJC-A1 DP  → C  (Duplex, KT)
 *   MOJC-SM    → E  (2.0mm, LG)
 *   OJC-C2-    → E  (2.0mm, KT)
 *   DROP-CABLE → K  (3.0mm)
 *   PIGTAIL-   → 규격에서 직경 파싱 후 F/G/H
 *   Optical Cable Parts → M
 */
export function detectOjcType(
  productName: string,
  spec: string = ""
): string | null {
  const n = productName.toUpperCase().trim();

  if (n.startsWith("SOJC-"))                            return "A";
  if (n.startsWith("OJC-A1-") && n.endsWith("-SP"))     return "A";
  if (n.startsWith("DOJC-"))                            return "C";
  if (n.startsWith("OJC-A1-") && n.endsWith("-DP"))     return "C";
  if (n.startsWith("OJC-A1-"))                          return "A"; // SP 기본값
  if (n.startsWith("MOJC-SM") || n.startsWith("OJC-C2-")) return "E";
  if (n.startsWith("DROP-CABLE"))                       return "K";
  if (n.startsWith("PIGTAIL-")) {
    // 규격에서 직경 파싱
    const s = spec.toUpperCase();
    if (s.includes("3.0MM") || s.includes("3.0 MM"))   return "H";
    if (s.includes("2.0MM") || s.includes("2.0 MM"))   return "G";
    if (s.includes("0.9MM") || s.includes("0.9 MM"))   return "F";
    return null; // 직경 정보 필요
  }
  if (n.startsWith("OPTICAL CABLE PARTS"))              return "M";
  return null;
}

/**
 * ② 코어 종류 코드 자동 판별
 *
 *   품목명에 -SM- 포함 → D (SMF 657.B3)   ← AJW 기준: SM = SMF 657.B3
 *   규격에 B3 / 657B3  → D
 *   규격에 A2 / 657A2  → C
 *   규격에 A1 / 657A1  → B
 *   규격에 652D        → A
 *   규격에 OM1~OM5     → E~I
 */
export function detectCoreType(
  productName: string,
  spec: string
): string | null {
  const n = productName.toUpperCase();
  const s = spec.toUpperCase();
  // 규격에 명시된 SMF 타입 최우선
  if (/657\.?B3/.test(s) || /\bB3\b/.test(s))  return "D";
  if (/657\.?A2/.test(s) || /\bA2\b/.test(s))  return "C";
  if (/657\.?A1/.test(s) || /\bA1\b/.test(s))  return "B";
  if (/652D?/.test(s))                            return "A";
  if (s.includes("OM5")) return "I";
  if (s.includes("OM4")) return "H";
  if (s.includes("OM3")) return "G";
  if (s.includes("OM2")) return "F";
  if (s.includes("OM1")) return "E";
  // 규격에 없으면 품목명 SM 표기 → SMF 657.B3
  if (/-SM[-_]/.test(n) || n.endsWith("-SM")) return "D";
  return null;
}

/**
 * ③ 심선수 코드 자동 판별
 *
 *   품목명 NC 패턴  → 코드 (예: 6C → "4", 12C → "8")
 *   DOJC           → "2" (항상 2심)
 *   SOJC / OJC-A1 SP → "1" (항상 1심)
 */
export function detectCoreCount(productName: string): string | null {
  const n = productName.toUpperCase();
  // -NC 패턴  (예: -6C, -12C)
  const m = n.match(/[-_](\d+)C(?=$|[-_\s])/);
  if (m) return CORE_COUNT_INV[parseInt(m[1])] ?? null;
  // NCore 패턴  (예: 2Core, 12Core — Optical Cable Parts 등)
  const mCore = n.match(/(\d+)CORE/);
  if (mCore) return CORE_COUNT_INV[parseInt(mCore[1])] ?? null;
  // 제품명 기본값
  if (n.startsWith("DOJC-"))                        return "2";
  if (n.startsWith("SOJC-"))                        return "1";
  if (n.startsWith("OJC-A1-") && n.endsWith("-SP")) return "1";
  if (n.startsWith("OJC-A1-") && n.endsWith("-DP")) return "2";
  // DROP-CABLE: 괄호 안 커넥터 2종이면 2심, 단일이면 1심
  if (n.startsWith("DROP-CABLE")) {
    const dm = n.match(/DROP-CABLE\(([^)]+)\)/);
    return (dm && dm[1].includes("-")) ? "2" : "1";
  }
  return null;
}

/**
 * ④ 재질 코드 자동 판별 — 코어 종류 코드 + 직경 + 제품 유형 기반
 *
 *   B3(D) + 2.0mm               → 0  (PVC YELLOW)
 *   B3(D) + 3.0mm               → 0  (PVC YELLOW)
 *   A1(B) + 2.0mm               → 0  (PVC YELLOW)
 *   A1(B) / A2(C) + 3.0mm DROP → B  (LSZH BLACK)
 *   MOJC / OJC-C2               → C  (LSZH YELLOW)
 *   Optical Cable Parts         → B  (LSZH BLACK)
 *   Pigtail                     → 규격에서 색수 파싱
 *                                    12색 → A, 6색 → 8, YELLOW → 0
 */
export function detectMaterial(
  productName: string,
  coreCode: string,
  spec: string,
  diameterMm: number | null
): string | null {
  const n = productName.toUpperCase();

  // MOJC / OJC-C2 → LSZH YELLOW
  if (n.startsWith("MOJC-") || n.startsWith("OJC-C2-")) return "C";

  // Optical Cable Parts → LSZH BLACK
  if (n.startsWith("OPTICAL CABLE PARTS")) return "B";

  // Pigtail → 규격에서 색수 파싱
  if (n.startsWith("PIGTAIL-")) {
    const s = spec.toUpperCase();
    if (s.includes("12색") || s.includes("12COLOR")) return "A";
    if (s.includes("6색")  || s.includes("6COLOR"))  return "8";
    return "0"; // 기본 YELLOW
  }

  // DROP-CABLE A1/A2 3.0mm → LSZH BLACK
  if (n.startsWith("DROP-CABLE") &&
      (coreCode === "B" || coreCode === "C") &&
      diameterMm === 3.0) return "B";

  // B3 + 2.0/3.0mm, A1 + 2.0mm → PVC YELLOW
  if (coreCode === "D") return "0";
  if (coreCode === "B" && diameterMm === 2.0) return "0";

  return null;
}

/**
 * ⑤⑥ 커넥터 A/B 코드 자동 판별 — 품목명 기반
 *
 *   DOJC/SOJC/MOJC: 품목명에서 직접 추출
 *     예) DOJC-SM-LC/PC-LC/PC → ["4","4"]
 *
 *   OJC-A1 / OJC-C2: pos3=커넥터쌍 + pos6=페롤쌍 조합
 *     예) OJC-A1-SC/LC-SM-3-APC/PC-SP → SC+APC="SC/APC"(1), LC+PC="LC/PC"(4)
 *
 *   DROP-CABLE: 괄호 내 추출
 *     예) DROP-CABLE(LC/PC-SC/APC) → ["4","1"]
 *     예) DROP-CABLE(LC/PC)        → ["4","4"]  (양끝 동일)
 *
 *   PIGTAIL: 커넥터A만 (B 없음)
 *     예) PIGTAIL-LC/APC-SM-12C → ["5", null]
 */
export function detectConnectors(
  productName: string
): [string | null, string | null] {
  const n = productName.toUpperCase().trim();

  // DROP-CABLE(XX) or DROP-CABLE(XX-YY)
  const dropM = n.match(/DROP-CABLE\(([^)]+)\)/);
  if (dropM) {
    const parts = dropM[1].split("-");
    const ca = CONNECTOR_ALIAS[parts[0].trim()] ?? null;
    const cb = parts[1]
      ? (CONNECTOR_ALIAS[parts[1].trim()] ?? null)
      : ca;  // 단일 표기 → 양끝 동일
    return [ca, cb];
  }

  // OJC-A1 / OJC-C2: 커넥터+페롤 분리 조합
  if (n.startsWith("OJC-A1-") || n.startsWith("OJC-C2-")) {
    const parts = n.split("-");
    // [OJC, A1/C2, SC/LC, SM, 길이, PC/PC, 심선수/SP/DP]
    if (parts.length >= 6) {
      const connPair   = parts[2].split("/");  // ["SC","LC"]
      const ferulePair = parts[5].split("/");  // ["APC","PC"]
      if (connPair.length === 2 && ferulePair.length === 2) {
        const ta = `${connPair[0]}/${ferulePair[0]}`;
        const tb = `${connPair[1]}/${ferulePair[1]}`;
        return [CONNECTOR_ALIAS[ta] ?? null, CONNECTOR_ALIAS[tb] ?? null];
      }
    }
  }

  // PIGTAIL: 커넥터A만
  const pigM = n.match(/PIGTAIL-((?:SC|LC|FC)\/(?:PC|APC))/);
  if (pigM) return [CONNECTOR_ALIAS[pigM[1]] ?? null, null];

  // DOJC / SOJC / MOJC / Optical Cable Parts: SC/PC 패턴 순서대로 추출
  const found = [...n.matchAll(/(SC|LC|FC|ST|MU)\/(PC|APC)(?:\s*CLIP)?/g)]
    .map(m => {
      const raw = m[0].replace(/\s+/g, " ").trim();
      return CONNECTOR_ALIAS[raw] ?? null;
    });
  if (found.length > 0) {
    const ca = found[0];
    // Simplex(SOJC/OJC-A1 SP): 커넥터 하나만 표기 → 양끝 동일
    const isSimplexName = n.startsWith("SOJC-") ||
      (n.startsWith("OJC-A1-") && n.endsWith("-SP"));
    const cb = found[1] ?? (isSimplexName ? ca : null);
    return [ca, cb];
  }

  return [null, null];
}

// ────────────────────────────────────────────────────────────────────────────
// 메인 파싱 & 품번 생성
// ────────────────────────────────────────────────────────────────────────────

/** 직경(mm) 추출 — 규격 문자열 또는 인자 */
function resolveDiameter(spec: string, hint?: number): number | null {
  if (hint) return hint;
  const s = spec.toUpperCase();
  const m = s.match(/(\d+(?:\.\d+)?)\s*MM(?!M)/);
  return m ? parseFloat(m[1]) : null;
}

/**
 * 품목명 + 규격 → 품번 구성 필드 자동 파싱
 *
 * @param productName  품목명  예) "DOJC-SM-LC/PC-LC/PC"
 * @param spec         규격    예) "17M, B3, 2.0mm"
 * @param diameterMm   직경(mm) 직접 지정 시 규격 파싱보다 우선
 * @param overrides    수동 입력 필드 (자동 판별 결과 덮어쓰기)
 */
export function parseProduct(
  productName: string,
  spec: string,
  diameterMm?: number,
  overrides?: Partial<PartFields>
): PartFields & { missing: string[] } {
  const missing: string[] = [];
  const diam = resolveDiameter(spec, diameterMm);

  const ojcType    = detectOjcType(productName, spec);
  const coreType   = detectCoreType(productName, spec);
  const coreCount  = detectCoreCount(productName);
  const material   = coreType
    ? detectMaterial(productName, coreType, spec, diam)
    : null;
  const [connectorA, connectorB] = detectConnectors(productName);
  const marking    = detectMarking(productName);

  const lengthM    = parseLength(spec);
  let lengthCode: string | null = null;
  if (lengthM !== null) {
    try { lengthCode = lengthToCode(lengthM); }
    catch { missing.push("lengthCode"); }
  } else {
    missing.push("lengthCode");
  }

  if (!ojcType)    missing.push("ojcType");
  if (!coreType)   missing.push("coreType");
  if (!coreCount)  missing.push("coreCount");
  if (!material)   missing.push("material");
  if (!connectorA) missing.push("connectorA");
  const isPigtail = productName.toUpperCase().startsWith("PIGTAIL");
  if (!connectorB && !isPigtail) missing.push("connectorB");

  const fields: PartFields = {
    ojcType, coreType, coreCount, material,
    connectorA, connectorB, lengthCode, marking,
    ...overrides,
  };

  // overrides로 해소된 missing 제거
  const resolvedMissing = missing.filter(k => {
    const key = k.replace(/\s.*/, "") as keyof PartFields;
    return !fields[key];
  });

  return { ...fields, missing: resolvedMissing };
}

/**
 * 품목명 + 규격 → 최종 품번 생성
 *
 * @returns GenerateResult
 *   - partNumber: 생성된 품번 (예: "A14-C-D20-44-17JL")
 *   - missing:    확인 필요 항목 목록
 *   - error:      오류 메시지 (생성 실패 시)
 */
export function generatePartNumber(
  productName: string,
  spec: string,
  diameterMm?: number,
  overrides?: Partial<PartFields>
): GenerateResult {
  const { missing, ...fields } = parseProduct(
    productName, spec, diameterMm, overrides
  );

  if (missing.length > 0) {
    return {
      partNumber: null,
      fields,
      missing,
      error: `확인 필요 항목: ${missing.join(", ")}`,
    };
  }

  const cb = fields.connectorB ?? "";
  const partNumber = [
    `A14-${fields.ojcType}`,
    `${fields.coreType}${fields.coreCount}${fields.material}`,
    `${fields.connectorA}${cb}`,
    `${fields.lengthCode}${fields.marking}`,
  ].join("-");

  return { partNumber, fields, missing: [], error: null };
}

// ────────────────────────────────────────────────────────────────────────────
// 중복 체크 / 등록 (Node.js 환경)
// ────────────────────────────────────────────────────────────────────────────

/**
 * 품번리스트 Excel에서 중복 여부 확인
 * 서버화 시 DB 조회로 교체 예정
 */
export async function checkDuplicate(
  partNumber: string,
  listPath: string
): Promise<boolean> {
  // eslint-disable-next-line @typescript-eslint/no-var-requires
  const XLSX = require("xlsx");
  const wb   = XLSX.readFile(listPath);
  const ws   = wb.Sheets[wb.SheetNames[0]];
  const rows: string[][] = XLSX.utils.sheet_to_json(ws, { header: 1 });
  return rows.slice(1).some(row => String(row[0]).trim() === partNumber.trim());
}

/**
 * 품번리스트 Excel에 신규 품번 등록
 * 서버화 시 DB INSERT로 교체 예정
 */
export async function registerPart(
  partNumber: string,
  productName: string,
  spec: string,
  note: string = "",
  listPath: string
): Promise<void> {
  // eslint-disable-next-line @typescript-eslint/no-var-requires
  const XLSX = require("xlsx");
  const wb   = XLSX.readFile(listPath);
  const ws   = wb.Sheets[wb.SheetNames[0]];
  const rows: unknown[][] = XLSX.utils.sheet_to_json(ws, { header: 1 });
  const today = new Date().toISOString().split("T")[0];
  rows.push([partNumber, productName, spec, note, today]);
  const newWs = XLSX.utils.aoa_to_sheet(rows);
  wb.Sheets[wb.SheetNames[0]] = newWs;
  XLSX.writeFile(wb, listPath);
}

// ────────────────────────────────────────────────────────────────────────────
// 테스트
// ────────────────────────────────────────────────────────────────────────────

if (require.main === module) {
  const cases: [string, string, number | undefined, string | null][] = [
    // [품목명, 규격, 직경(mm), 예상품번]
    ["DOJC-SM-LC/PC-LC/PC",
      "17M, B3, 2.0mm",            2.0,  "A14-C-D20-44-17JL"],
    ["DOJC-SM-SC/PC-FC/APC",
      "60M, B3, 2.0mm",            2.0,  "A14-C-D20-09-60JL"],
    ["MOJC-SM-6C-SC/PC-SC/PC",
      "40M",                       2.0,  "A14-E-D4C-00-40JL"],
    ["OJC-C2-SC/LC-SM-10-PC/PC-12C",
      "10M",                       2.0,  "A14-E-D8C-04-10JK"],
    ["Optical Cable Parts,LC/PC-SC/APC-Single mode-5m-2Core",
      "657A2, LSZH, 5m",           3.0,  "A14-M-C2B-41-005N"],
    ["SOJC-SM-LC/APC",
      "3M, B3, 2.0mm",             2.0,  "A14-A-D10-55-003L"],
    ["OJC-A1-SC/LC-SM-3-APC/PC-SP",
      "3M",                        2.0,  "A14-A-D10-14-003K"],
    ["DROP-CABLE(LC/PC-SC/APC)",
      "10M",                       3.0,  "A14-K-B2B-41-10JN"],
    ["PIGTAIL-LC/APC-SM-12C",
      "1M, 12색, 0.9mm",           0.9,  "A14-F-D8A-5-001N"],
  ];

  console.log("\n" + "=".repeat(72));
  console.log(
    "품목명".padEnd(42) + "생성 품번".padEnd(22) + "검증"
  );
  console.log("=".repeat(72));

  for (const [name, spec, diam, expected] of cases) {
    const r  = generatePartNumber(name, spec, diam);
    const pn = r.partNumber ?? `⚠  ${r.missing.join(", ")}`;
    const ok = expected == null ? ""
      : r.partNumber === expected ? "✅"
      : `❌  expected: ${expected}`;
    console.log(`${name.slice(0,40).padEnd(42)}${pn.padEnd(22)}${ok}`);
  }

  console.log("=".repeat(72) + "\n");
}
