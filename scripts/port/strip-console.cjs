// One-off migration tool: removes active console.log/console.warn debug
// statements left in the legacy rendering code.
const fs = require("fs");
const ts = require("typescript");
for (const file of process.argv.slice(2)) {
  const text = fs.readFileSync(file, "utf8");
  const sf = ts.createSourceFile(file, text, ts.ScriptTarget.Latest, true, ts.ScriptKind.TS);
  const ranges = [];
  const visit = (n) => {
    if (ts.isExpressionStatement(n) && ts.isCallExpression(n.expression)) {
      const c = n.expression.expression;
      if (ts.isPropertyAccessExpression(c) && ts.isIdentifier(c.expression) && c.expression.text === "console" && ["log", "warn", "info", "debug"].includes(c.name.text)) {
        ranges.push([n.getStart(sf), n.end]);
        return;
      }
    }
    ts.forEachChild(n, visit);
  };
  visit(sf);
  let out = text;
  for (const [a, b] of ranges.sort((x, y) => y[0] - x[0])) out = out.slice(0, a) + "/* debug log removed */" + out.slice(b);
  fs.writeFileSync(file, out);
  console.log(file, ranges.length);
}
