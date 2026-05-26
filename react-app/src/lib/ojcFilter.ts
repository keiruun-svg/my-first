function normalizeName(name: string): string {
  return name.trim().toUpperCase().replace(/[\s_]+/g, '-')
}

export const OJC_PREFIXES: Record<string, string[]> = {
  'KT OJC':              ['OJC-A1-', 'OJC-C2-'],
  'LG OJC':              ['SOJC-', 'DOJC-', 'MOJC-'],
  'DROP':                ['DROP-CABLE'],
  '피그테일':            ['PIGTAIL-'],
  'Optical Cable Parts': ['Optical Cable Parts'],
  'DX-MM':               ['DX-MM'],
}

const DIST_PREFIXES = ['Distribution-CABLE', 'DISTRIBUTION CABLE']

export const COLOR_MAP_OJC: Record<string, string> = {
  'KT OJC':              'DEEAF1',
  'LG OJC':              'E2EFDA',
  'DROP':                'FFF2CC',
  '피그테일':            'FCE4D6',
  'Optical Cable Parts': 'F4E6FF',
  'DX-MM':               'EDEDED',
}

export function classifyOjc(name: string): string | null {
  const n = normalizeName(name)
  for (const [label, prefixes] of Object.entries(OJC_PREFIXES)) {
    if (prefixes.some(p => n.startsWith(normalizeName(p)))) return label
  }
  return null
}

export function classifyDist(name: string): boolean {
  const n = normalizeName(name)
  return DIST_PREFIXES.some(p => n.startsWith(normalizeName(p)))
}

// 분류는 됐지만 이카운트/EMP 품명이 OJC 규칙에 안 맞아 누락될 수 있는 품목 감지
export function looksLikeOjc(name: string): boolean {
  const n = normalizeName(name)
  return n.includes('OJC') || n.includes('PIGTAIL') ||
         n.startsWith('DROP-CABLE') ||
         n.startsWith('SOJC-') || n.startsWith('DOJC-') || n.startsWith('MOJC-')
}

// KT-SP/DP/MOJC, LG-SP/DP/MOJC, DROP-1C/2C, 피그테일-SK 성단용 세분화
export function classifyOjcDetailed(name: string): string | null {
  const n = normalizeName(name)
  if (n.startsWith('OJC-A1-') || n.startsWith('OJC-C2-')) {
    if (n.includes('-SP-') || n.endsWith('-SP'))     return 'KT OJC-SP'
    if (n.includes('-DP-') || n.endsWith('-DP'))     return 'KT OJC-DP'
    if (n.includes('-MOJC-') || n.endsWith('-MOJC')) return 'KT OJC-MOJC'
    return 'KT OJC'
  }
  if (n.startsWith('SOJC-')) return 'LG OJC-SP'
  if (n.startsWith('DOJC-')) return 'LG OJC-DP'
  if (n.startsWith('MOJC-')) return 'LG OJC-MOJC'
  if (n.startsWith('DROP-CABLE')) {
    const core = n.split('-')[2] ?? ''
    if (core === '1C') return 'DROP-1C'
    if (core === '2C') return 'DROP-2C'
    return 'DROP'
  }
  if (n.startsWith('PIGTAIL-'))
    return (n.includes('성단') || n.includes('SK')) ? '피그테일-SK 성단용' : '피그테일'
  if (n.startsWith('OPTICAL-CABLE-PARTS')) return 'Optical Cable Parts'
  if (n.startsWith('DX-MM')) return 'DX-MM'
  return null
}

export const COLOR_MAP_OJC_DETAILED: Record<string, string> = {
  'KT OJC-SP':           'FFDEEAF1',
  'KT OJC-DP':           'FFBDD7EE',
  'KT OJC-MOJC':         'FFC5DEEF',
  'KT OJC':              'FFD9E8F5',
  'LG OJC-SP':           'FFE2EFDA',
  'LG OJC-DP':           'FFC6E0B4',
  'LG OJC-MOJC':         'FFD4EBCA',
  'DROP-1C':             'FFFFF2CC',
  'DROP-2C':             'FFFFE699',
  'DROP':                'FFFFD966',
  '피그테일-SK 성단용':   'FFFFD7C7',
  '피그테일':            'FFFCE4D6',
  'Optical Cable Parts': 'FFF4E6FF',
  'DX-MM':               'FFEDEDED',
}
