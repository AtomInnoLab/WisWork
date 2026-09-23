import { presentationPlanClaims } from '../../src/presentation-plan'
import type { PresentationPlan } from '../../src/presentation-plan'
import { benchmarkDeck } from './presentation-benchmark'

export function benchmarkPlan(): PresentationPlan {
  const deck = benchmarkDeck()
  return {
    version: 1,
    projectId: deck.id,
    title: deck.title,
    brief: {
      objective: '科研汇报',
      audience: '研究团队',
      language: 'zh-CN',
      minutes: 15,
      requiredContent: ['示例内容'],
      constraints: [],
    },
    sources: [
      {
        id: 'source',
        title: '合成基准',
        uri: deck.claims[0]!.source,
        locator: '第 1 页',
        excerpt: '示例数据仅用于测试',
      },
    ],
    claims: [
      {
        id: 'source-1',
        statement: deck.claims[0]!.text,
        type: 'assumption',
        sourceIds: ['source'],
        confidence: 'low',
        reviewStatus: 'needs_review',
      },
    ],
    style: deck.style,
    slides: deck.slides.map((slide) => ({
      id: slide.id,
      title: slide.title,
      purpose: '解释研究结果',
      claimIds: slide.claimIds!,
      layout: 'content',
      requiredAssets: [],
      acceptanceCriteria: ['可编辑文本'],
    })),
  }
}

export function benchmarkPlannedDeck() {
  return { ...benchmarkDeck(), claims: presentationPlanClaims(benchmarkPlan()) }
}
