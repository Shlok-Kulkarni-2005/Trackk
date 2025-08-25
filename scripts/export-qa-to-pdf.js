'use strict';

const fs = require('fs');
const path = require('path');
const puppeteer = require('puppeteer');

async function main() {
	const mdPath = path.resolve(__dirname, '../docs/trackopz-interview-qa.md');
	const outPdfPath = path.resolve(__dirname, '../docs/trackopz-interview-qa.pdf');

	if (!fs.existsSync(mdPath)) {
		console.error('Markdown file not found at:', mdPath);
		process.exit(1);
	}

	const markdown = fs.readFileSync(mdPath, 'utf8');
	// Dynamically import ESM-only 'marked'
	const { marked } = await import('marked');
	const bodyHtml = marked.parse(markdown, { headerIds: true, mangle: false });

	const html = `<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>TrackOPZ Interview Q&A</title>
  <style>
    :root { --text:#0f172a; --muted:#475569; --border:#e2e8f0; --accent:#0ea5e9; }
    html, body { margin:0; padding:0; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, 'Fira Sans', 'Droid Sans', 'Helvetica Neue', Arial, 'Segoe UI Emoji', 'Apple Color Emoji', 'Noto Color Emoji', sans-serif; color: var(--text); line-height: 1.5; }
    .container { max-width: 900px; margin: 0 auto; padding: 32px 28px; }
    h1, h2, h3 { line-height:1.25; }
    h1 { font-size: 28px; margin: 0 0 16px; }
    h2 { font-size: 22px; margin: 28px 0 12px; padding-top: 10px; border-top: 2px solid var(--border); }
    h3 { font-size: 18px; margin: 18px 0 8px; }
    p { margin: 6px 0; }
    ul { margin: 8px 0 12px 20px; }
    li { margin: 4px 0; }
    code { background: #f8fafc; border: 1px solid var(--border); padding: 1px 4px; border-radius: 4px; font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, 'Liberation Mono', 'Courier New', monospace; font-size: 90%; }
    pre code { display:block; padding: 12px; overflow-x:auto; }
    a { color: var(--accent); text-decoration: none; }
    a:hover { text-decoration: underline; }
    hr { border: none; border-top: 1px solid var(--border); margin: 20px 0; }
    .title { font-size: 30px; font-weight: 700; margin-bottom: 4px; }
    .subtitle { color: var(--muted); margin-bottom: 24px; }
    /* Avoid awkward page breaks */
    h2, h3 { page-break-after: avoid; page-break-inside: avoid; }
    ul, ol, p { page-break-inside: avoid; }
    /* Tables (if any) */
    table { border-collapse: collapse; width: 100%; margin: 8px 0 12px; }
    th, td { border: 1px solid var(--border); padding: 6px 8px; text-align: left; }
    th { background: #f8fafc; }
  </style>
</head>
<body>
  <div class="container">
    <div class="title">TrackOPZ Interview Q&A</div>
    <div class="subtitle">Neatly formatted study guide generated from repository context</div>
    ${bodyHtml}
  </div>
</body>
</html>`;

	let browser;
	const commonArgs = [
		'--no-sandbox',
		'--disable-setuid-sandbox',
		'--disable-gpu',
		'--no-zygote',
		'--disable-dev-shm-usage'
	];
	try {
		browser = await puppeteer.launch({ headless: true, args: commonArgs });
	} catch (e) {
		try {
			const execPath = puppeteer.executablePath && puppeteer.executablePath();
			if (!execPath) throw e;
			browser = await puppeteer.launch({ headless: true, executablePath: execPath, args: commonArgs });
		} catch (e2) {
			console.error('Puppeteer failed to launch, last error:', e2);
			throw e2;
		}
	}
	try {
		const page = await browser.newPage();
		await page.setContent(html, { waitUntil: 'load' });
		await page.emulateMediaType('screen');
		await page.pdf({
			path: outPdfPath,
			format: 'A4',
			margin: { top: '20mm', right: '16mm', bottom: '20mm', left: '16mm' },
			printBackground: true
		});
		console.log('✅ PDF generated at:', outPdfPath);
	} finally {
		await browser.close();
	}
}

main().catch((err) => {
	console.error('Failed to export PDF:', err);
	process.exit(1);
});


