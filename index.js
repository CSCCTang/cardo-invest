// cardo-invest 的 Host 半：把包内 skill 注册到 ctx.skills。
// 依据《dsh 插件开发标准（现行版）》5.1、6.1。
import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

export const name = 'cardo-invest'
export const inject = ['skills']

const skillDir = join(dirname(fileURLToPath(import.meta.url)), 'skills', 'cardo-invest')

function skillBody(raw) {
  // 去掉 YAML frontmatter：模型侧只接正文
  return raw.replace(/^---\r?\n[\s\S]*?\r?\n---\r?\n/, '')
}

export function apply(ctx) {
  ctx.skills.register({
    name: 'cardo-invest',
    description: "Cardo Invest 投资决策流水线：财务参数、匡算与财务评价、三方案比选、建议书成文。",
    whenToUse: "用户要基于《财务参数.xlsx》产出《项目投资决策建议书》（匡算、财务评价、三方案比选、10 章成文）时使用。",
    content: skillBody(readFileSync(join(skillDir, 'SKILL.md'), 'utf8')),
    resourceBase: { kind: 'directory', path: skillDir },
  })
}
