// WIXEL 서버 실행 진입점. 기본은 이 컴퓨터에서만 접속합니다.
// 외부 접속: HOST=0.0.0.0 TABULA_TOKEN=긴-무작위-암호 node server.js
// 서버 로직은 server/app.js, 외부 HTTP 중계 보안은 server/network.js에 있습니다.
import { resolve } from 'node:path';
import { realpathSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { networkInterfaces } from 'node:os';
import { createWixelServer, serverConfig } from './server/app.js';

function isMain() {
  try { return !!process.argv[1] && realpathSync(resolve(process.argv[1])) === realpathSync(fileURLToPath(import.meta.url)); }
  catch { return false; }
}

if (isMain()) {
  try {
    const config = serverConfig();
    const server = createWixelServer(config);
    server.on('error', (error) => {
      console.error(error.code === 'EADDRINUSE' ? '서버 포트가 사용 중입니다. PORT를 바꾸거나 기존 서버를 종료하세요.' : '서버를 시작하지 못했습니다. 주소와 포트를 확인하세요.');
      process.exitCode = 1;
    });
    server.listen(config.port, config.host, () => {
      const port = server.address().port;
      console.log(`WIXEL: http://localhost:${port}`);
      if (config.host === '0.0.0.0') {
        for (const addrs of Object.values(networkInterfaces())) {
          for (const address of addrs ?? []) if (address.family === 'IPv4' && !address.internal) console.log(`  다른 기기에서: http://${address.address}:${port}`);
        }
      }
      console.log(`  저장 폴더: ${config.data}${config.token ? ' (토큰 필요)' : ' (이 컴퓨터에서만)'}`);
    });
  } catch (error) {
    console.error(error.message);
    process.exitCode = 1;
  }
}
