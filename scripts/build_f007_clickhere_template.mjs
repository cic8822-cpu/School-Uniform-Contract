// F-007 HWPX의 고유 텍스트 토큰을 Kordoc 누름틀 필드로 전환한다.
import { readFile, writeFile } from 'node:fs/promises'

const [inputPath, outputPath] = process.argv.slice(2)
if (!inputPath || !outputPath) {
  throw new Error('사용법: node build_f007_clickhere_template.mjs <입력 section0.xml> <출력 section0.xml>')
}

const source = await readFile(inputPath, 'utf8')
let fieldId = 1700000101
const tokenNames = []
const result = source.replace(/<hp:t>([\s\S]*?)\{\{(F007[A-Z]+)\}\}([\s\S]*?)<\/hp:t>/g, (_match, before, name, after) => {
  tokenNames.push(name)
  const id = fieldId++
  const command = `Clickhere:set:48:Direction:wstring:${name.length}:${name} HelpState:wstring:0:  `
  return `<hp:t>${before}</hp:t><hp:ctrl><hp:fieldBegin id="${id}" type="CLICK_HERE" name="${name}" editable="1"><hp:parameters cnt="1" name=""><hp:stringParam name="Command">${command}</hp:stringParam></hp:parameters></hp:fieldBegin></hp:ctrl><hp:ctrl><hp:fieldEnd beginIDRef="${id}"/></hp:ctrl><hp:t>${after}</hp:t>`
})

const uniqueTokens = new Set(tokenNames)
if (uniqueTokens.size !== 21 || tokenNames.length !== 21 || /\{\{F007[A-Z]+\}\}/.test(result)) {
  throw new Error(`F-007 토큰 전환 실패: 치환=${tokenNames.length}, 고유=${uniqueTokens.size}`)
}

await writeFile(outputPath, result, 'utf8')
console.log(JSON.stringify({ replaced: tokenNames.length, fields: [...uniqueTokens] }))
