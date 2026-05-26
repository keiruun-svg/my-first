export interface LocationRule {
  keyword: string[]
  location: string
  description: string
}

export const LOCATION_RULES: LocationRule[] = [
  { keyword: ['drop-cable', 'drop optical cable', 'drop cable'], location: 'D1',    description: 'Drop Cable' },
  { keyword: ['dx-mm'],                                          location: 'D2-MM', description: 'DX-MM' },
  { keyword: ['dojc-mm'],                                        location: 'DO-MM', description: 'DOJC-MM' },
  { keyword: ['dojc'],                                           location: 'LG-DO', description: 'DOJC' },
  { keyword: ['sojc'],                                           location: 'LG-SO', description: 'SOJC' },
  { keyword: ['mojc'],                                           location: 'LG-MO', description: 'MOJC' },
  { keyword: ['adapter', '어댑터', '감쇠기', 'attenuator'],      location: 'EX',    description: '어댑터 / 감쇠기' },
  { keyword: ['열수축슬리브', 'splice protection sleeve'],        location: 'EX',    description: '열수축슬리브' },
  { keyword: ['ojc housing kit', 'housing kit'],                 location: 'M1',    description: 'Housing Kit' },
  { keyword: ['ferrule'],                                        location: 'M1',    description: 'Ferrule' },
  { keyword: ['pigtail'],                                        location: 'PI-TA', description: 'Pigtail' },
  { keyword: ['optical cable parts', 'optical cable part'],      location: 'O1',    description: 'Optical Cable Parts' },
  { keyword: ['optical cable 0.9'],                              location: 'K9-PI', description: 'Optical Cable 0.9mm' },
  { keyword: ['optical cable'],                                  location: 'K1',    description: 'Optical Cable' },
  { keyword: ['ojc'],                                            location: 'KT',    description: 'OJC' },
]

export function assignLocation(name: string): string | null {
  const n = name.toLowerCase()
  for (const rule of LOCATION_RULES) {
    if (rule.keyword.some(kw => n.includes(kw))) return rule.location
  }
  return null
}
