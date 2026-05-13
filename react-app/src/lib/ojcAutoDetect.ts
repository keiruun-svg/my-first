import type { OjcRules } from './types'

export interface DetectedFields {
  ojcType:    string | null
  coreType:   string | null
  coreCount:  string | null
  material:   string | null
  connectorA: string | null
  connectorB: string | null
  lengthCode: string | null
  marking:    string
}

const CONNECTOR_ALIAS: Record<string, string> = {
  'SC/PC': '0',       'SC/APC': '1',
  'SC/PC CLIP': '2',  'SC/PC,CLIP': '2',
  'SC/APC CLIP': '3', 'SC/APC,CLIP': '3',
  'LC/PC': '4',       'LC/APC': '5',
  'LC/PC CLIP': '6',  'LC/PC,CLIP': '6',
  'LC DUPLEX KIT(CLIP)': '6',
  'LC/APC CLIP': '7', 'LC/APC,CLIP': '7',
  'FC/PC': '8',       'FC/APC': '9',
  'ST/PC': 'A',       'ST/APC': 'B',
}

export function detectOjcType(name: string, spec: string): string | null {
  const n = name.toUpperCase().trim()
  if (n.startsWith('SOJC-'))                             return 'A'
  if (n.startsWith('OJC-A1-') && n.endsWith('-SP'))      return 'A'
  if (n.startsWith('DOJC-'))                             return 'C'
  if (n.startsWith('OJC-A1-') && n.endsWith('-DP'))      return 'C'
  if (n.startsWith('OJC-A1-'))                           return 'A'
  if (n.startsWith('MOJC-SM') || n.startsWith('OJC-C2-')) return 'E'
  if (n.startsWith('DROP-CABLE'))                        return 'K'
  if (n.startsWith('PIGTAIL-')) {
    const s = spec.toUpperCase()
    if (s.includes('3.0MM') || s.includes('3.0 MM')) return 'H'
    if (s.includes('2.0MM') || s.includes('2.0 MM')) return 'G'
    if (s.includes('0.9MM') || s.includes('0.9 MM')) return 'F'
    return null
  }
  if (n.startsWith('OPTICAL CABLE PARTS')) return 'M'
  return null
}

export function detectCoreType(name: string, spec: string): string | null {
  const s = spec.toUpperCase()
  const n = name.toUpperCase()
  if (/657\.?B3/.test(s) || /\bB3\b/.test(s)) return 'D'
  if (/657\.?A2/.test(s) || /\bA2\b/.test(s)) return 'C'
  if (/657\.?A1/.test(s) || /\bA1\b/.test(s)) return 'B'
  if (/652D?/.test(s))                         return 'A'
  if (s.includes('OM5')) return 'I'
  if (s.includes('OM4')) return 'H'
  if (s.includes('OM3')) return 'G'
  if (s.includes('OM2')) return 'F'
  if (s.includes('OM1')) return 'E'
  if (/-SM[-_]/.test(n) || n.endsWith('-SM')) return 'D'
  return null
}

export function detectCoreCount(name: string, rules: OjcRules): string | null {
  const n = name.toUpperCase()
  const countToCode: Record<number, string> = Object.fromEntries(
    rules.coreCount.map(e => [e.count, e.code])
  )
  const m = n.match(/[-_](\d+)C(?=$|[-_\s])/)
  if (m) return countToCode[parseInt(m[1])] ?? null
  const mCore = n.match(/(\d+)CORE/)
  if (mCore) return countToCode[parseInt(mCore[1])] ?? null
  if (n.startsWith('DOJC-'))                              return countToCode[2] ?? null
  if (n.startsWith('SOJC-'))                              return countToCode[1] ?? null
  if (n.startsWith('OJC-A1-') && n.endsWith('-SP'))       return countToCode[1] ?? null
  if (n.startsWith('OJC-A1-') && n.endsWith('-DP'))       return countToCode[2] ?? null
  if (n.startsWith('DROP-CABLE')) {
    const dm = n.match(/DROP-CABLE\(([^)]+)\)/)
    return (dm && dm[1].includes('-')) ? (countToCode[2] ?? null) : (countToCode[1] ?? null)
  }
  return null
}

export function detectMaterial(
  name: string,
  coreCode: string,
  spec: string,
  diamMm: number | null,
): string | null {
  const n = name.toUpperCase()
  if (n.startsWith('MOJC-') || n.startsWith('OJC-C2-'))  return 'C'
  if (n.startsWith('OPTICAL CABLE PARTS'))               return 'B'
  if (n.startsWith('PIGTAIL-')) {
    const s = spec.toUpperCase()
    if (s.includes('12색') || s.includes('12COLOR')) return 'A'
    if (s.includes('6색')  || s.includes('6COLOR'))  return '8'
    return '0'
  }
  if (n.startsWith('DROP-CABLE') && (coreCode === 'B' || coreCode === 'C') && diamMm === 3.0) return 'B'
  if (coreCode === 'D') return '0'
  if (coreCode === 'B' && diamMm === 2.0) return '0'
  return null
}

export function detectConnectors(name: string): [string | null, string | null] {
  const n = name.toUpperCase().trim()
  const dropM = n.match(/DROP-CABLE\(([^)]+)\)/)
  if (dropM) {
    const parts = dropM[1].split('-')
    const ca = CONNECTOR_ALIAS[parts[0].trim()] ?? null
    const cb = parts[1] ? (CONNECTOR_ALIAS[parts[1].trim()] ?? null) : ca
    return [ca, cb]
  }
  if (n.startsWith('OJC-A1-') || n.startsWith('OJC-C2-')) {
    const parts = n.split('-')
    if (parts.length >= 6) {
      const conn  = parts[2].split('/')
      const ferul = parts[5].split('/')
      if (conn.length === 2 && ferul.length === 2) {
        const ta = `${conn[0]}/${ferul[0]}`, tb = `${conn[1]}/${ferul[1]}`
        return [CONNECTOR_ALIAS[ta] ?? null, CONNECTOR_ALIAS[tb] ?? null]
      }
    }
  }
  const pigM = n.match(/PIGTAIL-((?:SC|LC|FC)\/(?:PC|APC))/)
  if (pigM) return [CONNECTOR_ALIAS[pigM[1]] ?? null, null]
  const found = [...n.matchAll(/(SC|LC|FC|ST|MU)\/(PC|APC)(?:\s*CLIP)?/g)]
    .map(m => CONNECTOR_ALIAS[m[0].replace(/\s+/g, ' ').trim()] ?? null)
  if (found.length > 0) {
    const isSimp = n.startsWith('SOJC-') || (n.startsWith('OJC-A1-') && n.endsWith('-SP'))
    return [found[0], found[1] ?? (isSimp ? found[0] : null)]
  }
  return [null, null]
}

export function detectMarking(name: string): string {
  const n = name.toUpperCase()
  if (n.startsWith('OJC-A1-') || n.startsWith('OJC-C2-'))         return 'K'
  if (n.startsWith('SOJC-') || n.startsWith('DOJC-') || n.startsWith('MOJC-')) return 'L'
  return 'N'
}

export function parseLength(spec: string): number | null {
  const s = spec.toUpperCase()
  const mm = s.match(/(\d+(?:\.\d+)?)\s*MM(?!\s*M)/)
  if (mm) return parseFloat(mm[1]) / 1000
  const m = s.match(/(\d+(?:\.\d+)?)\s*M(?!M)/)
  if (m) return parseFloat(m[1])
  return null
}

export function resolveLengthCode(m: number, presets: OjcRules['lengthPreset']): string | null {
  const preset = presets.find(p => p.m === m)
  if (preset) return preset.code
  if ([100, 150, 200, 250, 300].includes(m)) return String(m)
  const dec = (d: number) => String.fromCharCode(64 + d)
  if (m < 1) {
    const d = Math.round(m * 10)
    return (d >= 1 && d <= 9) ? `00${dec(d)}` : null
  }
  if (m < 10) {
    const i = Math.floor(m), d = Math.round((m - i) * 10)
    return d === 0 ? `00${i}` : `0${i}${dec(d)}`
  }
  if (m < 100) {
    const i = Math.floor(m), d = Math.round((m - i) * 10)
    return d === 0 ? `${String(i).padStart(2, '0')}J` : `${String(i).padStart(2, '0')}${dec(d)}`
  }
  return null
}

export function analyzeProduct(name: string, spec: string, diamMm: number | null, rules: OjcRules): DetectedFields {
  const ojcType    = detectOjcType(name, spec)
  const coreType   = detectCoreType(name, spec)
  const coreCount  = detectCoreCount(name, rules)
  const material   = coreType ? detectMaterial(name, coreType, spec, diamMm) : null
  const [connectorA, connectorB] = detectConnectors(name)
  const marking    = detectMarking(name)
  const lengthM    = parseLength(spec)
  const lengthCode = lengthM !== null ? resolveLengthCode(lengthM, rules.lengthPreset) : null
  return { ojcType, coreType, coreCount, material, connectorA, connectorB: connectorB ?? null, lengthCode, marking }
}
