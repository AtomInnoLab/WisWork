import type { PresentationDeck, SlideIRElement } from '../../src/presentation'

/** Eight-page 16:9 Chinese fixture; metrics are synthetic, never presented as researched facts. */
export function benchmarkDeck(): PresentationDeck {
  const title = (text: string): SlideIRElement => ({
    kind: 'text',
    id: 'title',
    x: 1,
    y: 1,
    w: 10,
    h: 1,
    text,
    fontSize: 32,
  })
  const body = (text: string): SlideIRElement => ({
    kind: 'text',
    id: 'body',
    x: 1,
    y: 2.5,
    w: 10,
    h: 3,
    text,
  })
  const titles = [
    '科研汇报',
    '目录',
    '研究图文',
    '研究流程',
    '方案对比',
    '实验表格',
    '数据图表',
    '总结',
  ]
  const deck: PresentationDeck = {
    version: 1,
    id: 'benchmark-eight',
    title: '八页可编辑基准',
    style: {
      fontFace: 'Microsoft YaHei',
      background: 'FFFFFF',
      textColor: '172033',
      accentColor: '2255AA',
    },
    assets: [
      {
        id: 'pixel',
        mime: 'image/png',
        width: 1,
        height: 1,
        base64:
          'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jB1sAAAAASUVORK5CYII=',
        source: 'Synthetic fixture',
      },
    ],
    claims: [
      {
        id: 'source-1',
        text: '示例数据仅用于测试',
        source: '研究报告（合成基准）',
        locator: '第 1 页',
      },
    ],
    slides: titles.map((text, index) => ({
      id: `slide-${index + 1}`,
      title: text,
      notes: '演讲备注',
      claimIds: ['source-1'],
      elements: [title(text), body('示例内容：原生文本可在 PowerPoint 中编辑。')],
    })),
  }
  deck.slides[2]!.elements[1] = {
    kind: 'image',
    id: 'image',
    x: 1,
    y: 2.5,
    w: 3,
    h: 3,
    assetId: 'pixel',
    fit: 'contain',
  }
  deck.slides[3]!.elements[1] = {
    kind: 'shape',
    id: 'step',
    x: 1,
    y: 2.5,
    w: 3,
    h: 2,
    shape: 'roundRect',
  }
  deck.slides[4]!.elements[1] = {
    kind: 'shape',
    id: 'comparison',
    x: 1,
    y: 2.5,
    w: 3,
    h: 2,
    shape: 'rect',
  }
  deck.slides[5]!.elements[1] = {
    kind: 'table',
    id: 'table',
    x: 1,
    y: 2.5,
    w: 8,
    h: 2,
    rows: [
      ['方案', '结果'],
      ['甲', '120'],
      ['乙', '90'],
    ],
  }
  deck.slides[6]!.elements[1] = {
    kind: 'chart',
    id: 'chart',
    x: 1,
    y: 2.5,
    w: 8,
    h: 3,
    chartType: 'bar',
    categories: ['甲', '乙'],
    series: [{ name: '示例', values: [120, 90] }],
  }
  return deck
}
