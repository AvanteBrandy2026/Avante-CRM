16:27:14.906 Restored build cache from previous deployment (9AyMS3AMYj4LW3p7aXuyP67LbU5g)
16:27:15.213 Running "vercel build"
16:27:15.238 Vercel CLI 53.3.2
16:27:15.856 Installing dependencies...
16:27:18.920 
16:27:18.921 up to date in 3s
16:27:18.921 
16:27:18.921 7 packages are looking for funding
16:27:18.922   run `npm fund` for details
16:27:18.956 Running "npm run build"
16:27:19.097 
16:27:19.097 > avante-crm@1.0.0 build
16:27:19.098 > vite build
16:27:19.098 
16:27:19.317 vite v5.4.21 building for production...
16:27:19.372 transforming...
16:27:19.478 ✓ 6 modules transformed.
16:27:19.480 x Build failed in 135ms
16:27:19.480 error during build:
16:27:19.480 [vite:esbuild] Transform failed with 1 error:
16:27:19.481 /vercel/path0/src/App.jsx:1073:2: ERROR: Unexpected ")"
16:27:19.481 file: /vercel/path0/src/App.jsx:1073:2
16:27:19.481 
16:27:19.481 Unexpected ")"
16:27:19.481 1071|      document.body
16:27:19.481 1072|    );
16:27:19.481 1073|    );
16:27:19.481    |    ^
16:27:19.481 1074|  }
16:27:19.481 1075|  
16:27:19.481 
16:27:19.482     at failureErrorWithLog (/vercel/path0/node_modules/esbuild/lib/main.js:1472:15)
16:27:19.482     at /vercel/path0/node_modules/esbuild/lib/main.js:755:50
16:27:19.482     at responseCallbacks.<computed> (/vercel/path0/node_modules/esbuild/lib/main.js:622:9)
16:27:19.482     at handleIncomingPacket (/vercel/path0/node_modules/esbuild/lib/main.js:677:12)
16:27:19.482     at Socket.readFromStdout (/vercel/path0/node_modules/esbuild/lib/main.js:600:7)
16:27:19.482     at Socket.emit (node:events:509:28)
16:27:19.482     at addChunk (node:internal/streams/readable:563:12)
16:27:19.482     at readableAddChunkPushByteMode (node:internal/streams/readable:514:3)
16:27:19.482     at Readable.push (node:internal/streams/readable:394:5)
16:27:19.483     at Pipe.onStreamRead (node:internal/stream_base_commons:189:23)
16:27:19.498 Error: Command "npm run build" exited with 1
