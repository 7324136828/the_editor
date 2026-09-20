import { readFile } from 'node:fs/promises';
import ts from 'typescript';

const compiled = new Map();
/** Load the production TypeScript modules in Node without an extra test runtime. */
export async function importTypeScript(url) {
  async function compile(sourceUrl) {
    if (compiled.has(sourceUrl.href)) return compiled.get(sourceUrl.href);
    const source = await readFile(sourceUrl, 'utf8');
    let code = ts.transpileModule(source, { compilerOptions: { target: ts.ScriptTarget.ES2022, module: ts.ModuleKind.ES2022 } }).outputText;
    for (const match of [...code.matchAll(/from ['"](\.[^'"]+)['"]/g)]) {
      const dependency = await compile(new URL(`${match[1]}.ts`, sourceUrl));
      code = code.replace(match[0], `from '${dependency}'`);
    }
    const result = `data:text/javascript;base64,${Buffer.from(code).toString('base64')}`;
    compiled.set(sourceUrl.href, result);
    return result;
  }
  return import(await compile(url));
}
