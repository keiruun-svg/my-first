export const OJC_PREFIXES: Record<string, string[]> = {
  'KT OJC': ['OJC-A1-', 'OJC-C2-'],
  'LG OJC': ['SOJC-', 'DOJC-', 'MOJC-'],
  'DROP': ['DROP-CABLE'],
  '피그테일': ['PIGTAIL-'],
  'Optical Cable Parts': ['Optical Cable Parts'],
  'DX-MM': ['DX-MM'],
}

const DIST_PREFIXES = ['Distribution-CABLE', 'DISTRIBUTION CABLE']

export const COLOR_MAP_OJC: Record<string, string> = {
  'KT OJC': 'DEEAF1',
  'LG OJC': 'E2EFDA',
  'DROP': 'FFF2CC',
  '피그테일': 'FCE4D6',
  'Optical Cable Parts': 'F4E6FF',
  'DX-MM': 'EDEDED',
}

export function classifyOjc(name: string): string | null {
  for (const [label, prefixes] of Object.entries(OJC_PREFIXES)) {
    if (prefixes.some(p => name.startsWith(p))) return label
  }
  return null
}

export function classifyDist(name: string): boolean {
  return DIST_PREFIXES.some(p => name.startsWith(p))
}

// KT-SP/DP/MOJC, LG-SP/DP/MOJC, DROP-1C/2C, 피그테일-SK 성단용 세분화
export function classifyOjcDetailed(name: string): string | null {
  if (name.startsWith('OJC-A1-') || name.startsWith('OJC-C2-')) {
    const u = name.toUpperCase()
    if (u.includes('-SP-') || u.endsWith('-SP'))   return 'KT OJC-SP'
    if (u.includes('-DP-') || u.endsWith('-DP'))   return 'KT OJC-DP'
    if (u.includes('-MOJC-') || u.endsWith('-MOJC')) return 'KT OJC-MOJC'
    return 'KT OJC'
  }
  if (name.startsWith('SOJC-')) return 'LG OJC-SP'
  if (name.startsWith('DOJC-')) return 'LG OJC-DP'
  if (name.startsWith('MOJC-')) return 'LG OJC-MOJC'
  if (name.startsWith('DROP-CABLE')) {
    const core = (name.split('-')[2] ?? '').toUpperCase()
    if (core === '1C') return 'DROP-1C'
    if (core === '2C') return 'DROP-2C'
    return 'DROP'
  }
  if (name.startsWith('PIGTAIL-'))
    return (name.includes('성단') || name.toUpperCase().includes('SK')) ? '피그테일-SK 성단용' : '피그테일'
  if (name.startsWith('Optical Cable Parts')) return 'Optical Cable Parts'
  if (name.startsWith('DX-MM')) return 'DX-MM'
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
