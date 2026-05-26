import { defineConfig, Plugin, ViteDevServer } from 'vite';
import { resolve } from 'path';
import { spawn, ChildProcess } from 'child_process';
import { existsSync } from 'fs';
import http from 'http';

// ---------------------------------------------------------------------------
// Helper — poll Flask's /health until it responds 200 or we time out
// ---------------------------------------------------------------------------
function waitForFlask(timeoutMs = 15_000): Promise<void> {
  return new Promise((res, rej) => {
    const deadline = Date.now() + timeoutMs;
    const check = () =>
      http
        .get('http://127.0.0.1:5000/health', (r) => {
          if (r.statusCode === 200) { res(); return; }
          schedule();
        })
        .on('error', schedule);
    const schedule = () =>
      Date.now() < deadline
        ? setTimeout(check, 400)
        : rej(new Error('Flask did not start within the timeout period.'));
    check();
  });
}

// ---------------------------------------------------------------------------
// Vite plugin — auto-starts Flask, watches Flask source files, and triggers
// live reloads in the preview window whenever anything changes.
//
//   templates/*.html  →  full browser reload (Flask re-renders the template)
//   *.py              →  Flask process restart, then full browser reload
// ---------------------------------------------------------------------------
function flaskPlugin(): Plugin {
  let flask: ChildProcess | null = null;
  let python = 'python';
  let restarting = false;

  // Spawn Flask and pipe its output to the Vite console
  function spawnFlask(): ChildProcess {
    const proc = spawn(python, ['app.py'], {
      cwd: __dirname,
      stdio: ['ignore', 'pipe', 'pipe'],
    });
    proc.stdout?.on('data', (d: Buffer) =>
      process.stdout.write(`  \x1b[2m[flask]\x1b[0m ${d}`));
    proc.stderr?.on('data', (d: Buffer) =>
      process.stderr.write(`  \x1b[2m[flask]\x1b[0m ${d}`));
    proc.on('error', (err) =>
      console.error(`\n  \x1b[31m✗\x1b[0m  Flask process error: ${err.message}\n`));
    return proc;
  }

  // Kill current Flask process and start a fresh one, then reload browser
  async function restartFlask(server: ViteDevServer, changedFile: string) {
    if (restarting) return;
    restarting = true;
    const short = changedFile.replace(__dirname, '').replace(/\\/g, '/');
    console.log(`\n  \x1b[33m⟳\x1b[0m  ${short} changed — restarting Flask…\n`);
    if (flask) { flask.kill(); flask = null; }
    // Brief pause to let the port clear
    await new Promise(r => setTimeout(r, 400));
    flask = spawnFlask();
    try {
      await waitForFlask(15_000);
      console.log(`\n  \x1b[32m✓\x1b[0m  Flask restarted — reloading preview…\n`);
      server.ws.send({ type: 'full-reload', path: '*' });
    } catch {
      console.error(`\n  \x1b[31m✗\x1b[0m  Flask did not restart within 15 s.\n`);
    }
    restarting = false;
  }

  return {
    name: 'vite-plugin-flask',

    async configureServer(server) {
      // Resolve Python path once
      const venvWin  = resolve(__dirname, 'venv/Scripts/python.exe');
      const venvUnix = resolve(__dirname, 'venv/bin/python');
      python = existsSync(venvWin)  ? venvWin
             : existsSync(venvUnix) ? venvUnix
             : 'python';

      // ── Start Flask if not already running ────────────────────────────────
      const alreadyUp = await new Promise<boolean>((resolve) => {
        http
          .get('http://127.0.0.1:5000/health', (r) => resolve(r.statusCode === 200))
          .on('error', () => resolve(false));
      });

      if (alreadyUp) {
        console.log('\n  \x1b[32m✓\x1b[0m  Flask already running on port 5000\n');
      } else {
        console.log('\n  \x1b[33m⟳\x1b[0m  Starting Flask on port 5000…\n');
        flask = spawnFlask();
        try {
          await waitForFlask(15_000);
          console.log('\n  \x1b[32m✓\x1b[0m  Flask ready — preview at http://localhost:5173\n');
        } catch {
          console.error(
            '\n  \x1b[31m✗\x1b[0m  Flask did not start within 15 s.\n' +
            '     Check that the venv exists (run deploy.bat) and port 5000 is free.\n',
          );
        }
      }

      // ── Watch Flask source files for live reload ───────────────────────────
      // Extend Vite's watcher to cover files outside frontend/
      server.watcher.add([
        resolve(__dirname, 'templates/**/*.html'),  // Jinja2 templates
        resolve(__dirname, '*.py'),                  // Flask + Python modules
      ]);

      server.watcher.on('change', async (file) => {
        const isTemplate = file.endsWith('.html');
        const isPython   = file.endsWith('.py');

        if (isTemplate) {
          // Template change: Flask doesn't need restart, just reload the browser
          const short = file.replace(__dirname, '').replace(/\\/g, '/');
          console.log(`\n  \x1b[36m↻\x1b[0m  ${short} changed — reloading preview…\n`);
          server.ws.send({ type: 'full-reload', path: '*' });
        } else if (isPython) {
          // Python change: must restart Flask so new code is loaded
          await restartFlask(server, file);
        }
      });

      // ── Stop Flask when Vite shuts down ───────────────────────────────────
      server.httpServer?.once('close', () => {
        if (flask) {
          flask.kill();
          flask = null;
          console.log('\n  \x1b[2m[flask]\x1b[0m  Stopped.\n');
        }
      });
    },
  };
}

// ---------------------------------------------------------------------------
// Vite config
// ---------------------------------------------------------------------------
export default defineConfig({
  root: 'frontend',

  plugins: [flaskPlugin()],

  build: {
    outDir: resolve(__dirname, 'static/dist'),
    emptyOutDir: true,
    rollupOptions: {
      input: { main: resolve(__dirname, 'frontend/main.ts') },
      output: {
        entryFileNames: '[name].js',
        chunkFileNames: '[name].js',
        assetFileNames: '[name][extname]',
      },
    },
  },

  server: {
    port: 5173,
    strictPort: true,

    proxy: {
      '^/(?!@vite|@fs|node_modules)': {
        target: 'http://127.0.0.1:5000',
        changeOrigin: true,
        bypass(req) {
          const url = req.url ?? '';
          if (url.startsWith('/@') || url.startsWith('/node_modules')) return url;
        },
      },
    },
  },
});
