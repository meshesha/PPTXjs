// One-off migration tool: adds an explicit `: any` to every untyped `var`
// declaration in the ported legacy modules (the code relies on JS `var`
// semantics and dynamic shapes from the XML tree).
//
// usage: node scripts/port/annotate-vars.cjs src/render/*.ts
const fs = require("fs");
const ts = require("typescript");

for (const file of process.argv.slice(2)) {
  const text = fs.readFileSync(file, "utf8");
  const sf = ts.createSourceFile(file, text, ts.ScriptTarget.Latest, true, ts.ScriptKind.TS);
  const inserts = [];
  const visit = (node) => {
    if (ts.isVariableDeclarationList(node) && !(node.flags & (ts.NodeFlags.Let | ts.NodeFlags.Const))) {
      const parent = node.parent;
      const isForInOf = parent && (ts.isForInStatement(parent) || ts.isForOfStatement(parent));
      if (!isForInOf) {
        for (const d of node.declarations) {
          if (!d.type && ts.isIdentifier(d.name)) inserts.push(d.name.end);
        }
      }
    }
    if ((ts.isArrowFunction(node) || ts.isFunctionExpression(node)) && node.parameters) {
      for (const p of node.parameters) {
        if (!p.type && ts.isIdentifier(p.name)) inserts.push({ pos: p.name.end, text: "?: any" });
      }
    }
    ts.forEachChild(node, visit);
  };
  visit(sf);
  let out = text;
  const items = inserts.map((i) => (typeof i === "number" ? { pos: i, text: ": any" } : i)).sort((a, b) => b.pos - a.pos);
  for (const { pos, text: t } of items) out = out.slice(0, pos) + t + out.slice(pos);
  fs.writeFileSync(file, out);
  console.log(file, items.length, "annotations");
}
