import http from 'node:http';
import fs from 'node:fs';
const file = new URL('../qa/native-audio/01 - ÄÖÜ тест.wav', import.meta.url);
const server = http.createServer((request,response)=>{
 const size=fs.statSync(file).size;
 const match=/^bytes=(\d+)-(\d*)$/.exec(request.headers.range||'');
 let first=match?Number(match[1]):0,last=match&&match[2]?Number(match[2]):size-1;
 if(first>=size||last>=size){response.writeHead(416,{'Content-Range':`bytes */${size}`});response.end();return;}
 const headers={'Content-Type':'audio/wav','Accept-Ranges':'bytes','Content-Length':last-first+1};
 if(match)headers['Content-Range']=`bytes ${first}-${last}/${size}`;
 response.writeHead(match?206:200,headers);
 if(request.method==='HEAD')response.end();else fs.createReadStream(file,{start:first,end:last}).pipe(response);
});
server.listen(52751,'127.0.0.1',()=>console.log('Local audio fixture server on 127.0.0.1:52751'));
