import fs from 'node:fs';
const qa=new URL('../qa/',import.meta.url);
for(const d of ['audio/','native-audio/','logic/'])fs.mkdirSync(new URL(d,qa),{recursive:true});
for(const [name,seconds] of [['01 - ÄÖÜ тест.wav',30],['02 - Second.wav',35]]) {
 const n=seconds*44100*2,b=Buffer.alloc(44+n);
 b.write('RIFF');b.writeUInt32LE(36+n,4);b.write('WAVEfmt ',8);b.writeUInt32LE(16,16);
 b.writeUInt16LE(1,20);b.writeUInt16LE(1,22);b.writeUInt32LE(44100,24);b.writeUInt32LE(88200,28);
 b.writeUInt16LE(2,32);b.writeUInt16LE(16,34);b.write('data',36);b.writeUInt32LE(n,40);
 for(const d of ['audio/','native-audio/']){const file=new URL(d+name,qa);if(!fs.existsSync(file))fs.writeFileSync(file,b);}
}
console.log('Silent audio fixtures prepared');
