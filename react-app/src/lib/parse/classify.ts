export interface ProductInfo {
  kind:   string        // 'a1' | 'b3' | 'om1' | 'om3' | 'drop' | 'pigtail' | 'om1-pigtail' | 'a2' | 'b3-Nc' | 'a1-Nc' | 'a1-청' | ...
  pai:    string        // '2.0mm' | '3.0mm' | '0.9mm'
  core:   number
  length: number | null
  type1:  string        // 'LC/PC' | 'SC/APC' | ...
  type2:  string
}

export function extractTypes(name: string): [string, string] {
  name = name.trim()
  if (name.startsWith('DROP')) {
    const m = name.match(/\((.+?)\)/)
    if (m) {
      const t = m[1].match(/(?:SC|LC|FC)\/(?:PC|APC)/g) ?? []
      if (t.length >= 2) return [t[0]!, t[1]!]
      if (t.length === 1) return [t[0]!, t[0]!]
    }
    return ['', '']
  }
  if (name.startsWith('PIGTAIL')) {
    const t = name.match(/(?:SC|LC|FC)\/(?:PC|APC)/g) ?? []
    return t.length ? [t[0]!, ''] : ['', '']
  }
  if (name.startsWith('OJC-')) {
    const parts = name.split('-')
    const conn  = parts.find(x => /^(SC|LC|FC)\/(SC|LC|FC)$/.test(x))
    const ferr  = parts.find(x => /^(PC|APC)\/(PC|APC)$/.test(x))
    if (conn && ferr) {
      const [cA, cB] = conn.split('/'); const [fA, fB] = ferr.split('/')
      return [`${cA}/${fA}`, `${cB}/${fB}`]
    }
    return ['', '']
  }
  if (/^(?:SOJC|DOJC|MOJC|Optical Cable Parts)/.test(name)) {
    const t = name.match(/(?:SC|LC|FC)\/(?:PC|APC)/g) ?? []
    if (t.length >= 2) return [t[0]!, t[1]!]
    if (t.length === 1) return [t[0]!, t[0]!]
    return ['', '']
  }
  return ['', '']
}

export function deriveCore(name: string): number {
  let m: RegExpMatchArray | null
  m = name.match(/MOJC-(?:SM|MM)-(\d+)C/);         if (m) return parseInt(m[1])
  m = name.match(/PIGTAIL-[A-Z/()A-Z-]+-(\d+)C\b/); if (m) return parseInt(m[1])
  if (name.startsWith('PIGTAIL')) { m = name.match(/-(\d+)C\b/); return m ? parseInt(m[1]) : 1 }
  m = name.match(/OJC-C2-.*-(\d+)C/);               if (m) return parseInt(m[1])
  if (name.startsWith('Optical Cable Parts')) { m = name.match(/(\d+)Core/); return m ? parseInt(m[1]) : 1 }
  if (name.startsWith('DROP') && name.includes('2C')) return 2
  if (name.startsWith('DROP'))  return 1
  if (name.startsWith('DOJC')) return 2
  if (name.startsWith('SOJC')) return 1
  if (name.endsWith('-SP')) return 1
  if (name.endsWith('-DP')) return 2
  m = name.match(/-(\d+)C$/); if (m) return parseInt(m[1])
  return 1
}

export function deriveKind(name: string, spec: string, _core: number): string {
  if (name.startsWith('PIGTAIL'))
    return /-MM\b|-MM-|MM\(OM3\)/.test(name) ? 'om1-pigtail' : 'pigtail'
  if (name.startsWith('DROP'))                return 'drop'
  if (name.startsWith('Optical Cable Parts')) return 'a2'
  if (name.startsWith('MOJC'))                return 'b3'
  if (/\bB3\b/i.test(spec))                   return 'b3'
  if (name.includes('MM(OM3)') || /\bOM3\b/i.test(spec) || name.includes('-OM3')) return 'om3'
  if (/-MM\b|-MM-/.test(name))                return 'om1'
  // 색상 — 규격 우선, 그 다음 품명
  for (const [c, l] of [['청','a1-청'],['적','a1-적'],['녹','a1-녹'],['자','a1-자']] as const)
    if (spec.includes(c)) return l
  for (const [c, l] of [['청','a1-청'],['적','a1-적'],['녹','a1-녹'],['자','a1-자']] as const)
    if (name.includes(c)) return l
  return 'a1'
}

export function derivePai(name: string, spec: string): string {
  if (name.startsWith('PIGTAIL')) {
    if (name.includes('(2.0mm)') || /-MM\b|-MM-/.test(name) || name.includes('MM(OM3)')) return '2.0mm'
    return '0.9mm'
  }
  if (name.startsWith('DROP')) return '3.0mm'
  if (/-MM\b|-MM-/.test(name) || name.includes('MM(OM3)') || name.includes('OM3')) return '2.0mm'
  if (/^(?:MOJC|Optical Cable Parts)/.test(name) || name.includes('OJC-C2') || name.includes('OJC-A1')) return '2.0mm'
  if (/^(?:SOJC|DOJC)/.test(name)) return /3\.0mm|3MM/i.test(spec) ? '3.0mm' : '2.0mm'
  if (/3\.0mm|3\.0MM/i.test(spec)) return '3.0mm'
  return '2.0mm'
}

export function deriveLength(name: string, spec: string): number | null {
  let m: RegExpMatchArray | null
  if (name.startsWith('Optical Cable Parts')) {
    m = name.match(/-(\d+(?:\.\d+)?)m-/i); if (m) return parseFloat(m[1])
  }
  if (name.includes('OJC-A1') || name.includes('OJC-C2')) {
    m = name.match(/-(?:SM|MM)(?:\(OM3\))?-(\d+(?:\.\d+)?)-(?:PC|APC)/); if (m) return parseFloat(m[1])
  }
  m = spec.match(/(\d+(?:\.\d+)?)\s*[mM]\b/); if (m) return parseFloat(m[1])
  m = spec.match(/(\d+(?:\.\d+)?)\s*[mM]/);   if (m) return parseFloat(m[1])
  m = name.match(/\[(\d+(?:\.\d+)?)M/i);      if (m) return parseFloat(m[1])
  return null
}

export function classifyRow(name: string, spec: string): ProductInfo {
  const core = deriveCore(name)
  return {
    kind:   deriveKind(name, spec, core),
    pai:    derivePai(name, spec),
    core,
    length: deriveLength(name, spec),
    type1:  extractTypes(name)[0],
    type2:  extractTypes(name)[1],
  }
}
