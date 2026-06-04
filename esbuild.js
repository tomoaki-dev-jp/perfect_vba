// Build script for the Perfect VBA extension.
// Bundles src/extension.ts -> dist/extension.js and copies the PowerShell
// bridge scripts to dist/scripts so they resolve relative to the bundle.
const esbuild = require("esbuild");
const fs = require("fs");
const path = require("path");

const production = process.argv.includes("--production");
const watch = process.argv.includes("--watch");

const UTF8_BOM = Buffer.from([0xef, 0xbb, 0xbf]);

/**
 * .ps1 を UTF-8 BOM 付きで書き出す。
 * Windows PowerShell 5.1 は BOM 無しのスクリプトをシステム ANSI(日本語環境では
 * cp932)として読むため、日本語コメント/文字列を含むと文字化け・構文破壊を起こす。
 * BOM を付けると UTF-8 として正しく解釈される。
 */
function copyPs1WithBom(src, dest) {
  let buf = fs.readFileSync(src);
  if (!(buf.length >= 3 && buf[0] === 0xef && buf[1] === 0xbb && buf[2] === 0xbf)) {
    buf = Buffer.concat([UTF8_BOM, buf]);
  }
  fs.writeFileSync(dest, buf);
}

/** Recursively copy a directory (only the files matching `filter`). */
function copyDir(src, dest, filter) {
  if (!fs.existsSync(src)) return;
  fs.mkdirSync(dest, { recursive: true });
  for (const entry of fs.readdirSync(src, { withFileTypes: true })) {
    const s = path.join(src, entry.name);
    const d = path.join(dest, entry.name);
    if (entry.isDirectory()) {
      copyDir(s, d, filter);
    } else if (!filter || filter(entry.name)) {
      if (entry.name.endsWith(".ps1")) {
        copyPs1WithBom(s, d);
      } else {
        fs.copyFileSync(s, d);
      }
    }
  }
}

function copyAssets() {
  copyDir(
    path.join(__dirname, "src", "bridge", "scripts"),
    path.join(__dirname, "dist", "scripts"),
    (name) => name.endsWith(".ps1")
  );
}

/** esbuild plugin that recopies the .ps1 assets after every build. */
const copyAssetsPlugin = {
  name: "copy-assets",
  setup(build) {
    build.onEnd(() => copyAssets());
  },
};

async function main() {
  const ctx = await esbuild.context({
    entryPoints: ["src/extension.ts"],
    bundle: true,
    format: "cjs",
    platform: "node",
    target: "node18",
    outfile: "dist/extension.js",
    external: ["vscode"],
    sourcemap: !production,
    minify: production,
    logLevel: "info",
    plugins: [copyAssetsPlugin],
  });

  if (watch) {
    await ctx.watch();
  } else {
    await ctx.rebuild();
    await ctx.dispose();
  }
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
