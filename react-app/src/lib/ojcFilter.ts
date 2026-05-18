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
