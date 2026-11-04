const {test}=require('node:test');
const assert=require('node:assert/strict');
const {buildDocx,wordFilename}=require('../word-export.js');

// Read the produced ZIP at its public file boundary, independently of writer helpers.
function parts(bytes){
 const data=Buffer.from(bytes),files=new Map();let offset=0;
 while(data.readUInt32LE(offset)===0x04034b50){
  const method=data.readUInt16LE(offset+8),size=data.readUInt32LE(offset+18),nameLength=data.readUInt16LE(offset+26),extraLength=data.readUInt16LE(offset+28);
  assert.equal(method,0,'parts use stored ZIP entries');
  const name=data.subarray(offset+30,offset+30+nameLength).toString('utf8');
  const start=offset+30+nameLength+extraLength;
  files.set(name,data.subarray(start,start+size).toString('utf8'));offset=start+size;
 }
 assert.equal(data.readUInt32LE(offset),0x02014b50,'central directory follows the local entries');
 const eocd=data.length-22;assert.equal(data.readUInt32LE(eocd),0x06054b50);assert.equal(data.readUInt32LE(eocd+16),offset);
 assert.equal(data.readUInt16LE(eocd+10),files.size);
 return files;
}
function paragraphs(xml){
 const decode=s=>s.replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/&quot;/g,'"').replace(/&apos;/g,"'").replace(/&amp;/g,'&');
 return [...xml.matchAll(/<w:p(?:\s[^>]*)?>([\s\S]*?)<\/w:p>/g)].map(m=>[...m[1].matchAll(/<w:t\b[^>]*>([\s\S]*?)<\/w:t>|<w:tab\s*\/>/g)].map(x=>x[1]===undefined?'\t':decode(x[1])).join(''));
}
test('exports a genuine ZIP-based editable DOCX with required relationships',()=>{
 const p=parts(buildDocx('周会\n\n请校长讲话。'));
 for(const name of ['[Content_Types].xml','_rels/.rels','word/document.xml','word/styles.xml','word/_rels/document.xml.rels'])assert.ok(p.has(name),name);
 assert.match(p.get('[Content_Types].xml'),/wordprocessingml.document.main\+xml/);
 assert.match(p.get('_rels/.rels'),/Target="word\/document.xml"/);
 assert.match(p.get('word/document.xml'),/<w:t[^>]*>请校长讲话。<\/w:t>/);
});
test('preserves every manual line, blank, space, tab and XML special character',()=>{
 const source='\n2026年11月11日周会\n\n  手动补充 <留步> & A > B "确认"\t老师  \n测试同学 🐡\n';
 const xml=parts(buildDocx(source)).get('word/document.xml');
 assert.deepEqual(paragraphs(xml),source.split('\n'));assert.match(xml,/xml:space="preserve"/);
});
test('normalizes Windows line endings without merging or discarding paragraphs',()=>{
 assert.deepEqual(paragraphs(parts(buildDocx('标题\r\n\r\n第一位\r第二位')).get('word/document.xml')),['标题','','第一位','第二位']);
});
test('drops only XML-forbidden controls while preserving all valid text',()=>{
 assert.deepEqual(paragraphs(parts(buildDocx('标题\n甲\u0000乙\u000b丙\t丁')).get('word/document.xml')),['标题','甲乙丙\t丁']);
});
test('rejects empty or whitespace-only scripts',()=>{
 for(const source of ['', ' \n\t '])assert.throws(()=>buildDocx(source),/司仪稿|内容/);
});
test('keeps a competition introduction and first recipient together',()=>{
 const xml=parts(buildDocx('周会\n\n在演讲比赛中，\n高一文忠 测试同学 荣获第二名\n请获奖同学与老师一同合影留念。')).get('word/document.xml');
 const ps=[...xml.matchAll(/<w:p(?:\s[^>]*)?>([\s\S]*?)<\/w:p>/g)].map(x=>x[1]);
 assert.match(ps[2],/<w:keepNext\s*\/>/);assert.match(ps[3],/<w:keepNext\s*\/>/);
});
test('does not force an entire long list to stay on one page',()=>{
 const source='周会\n\n在团体比赛中，代表同学有\n'+Array.from({length:100},(_,i)=>'高一文忠 同学'+i).join('\n');
 const xml=parts(buildDocx(source)).get('word/document.xml');
 const ps=[...xml.matchAll(/<w:p(?:\s[^>]*)?>([\s\S]*?)<\/w:p>/g)].map(x=>x[1]);
 assert.match(ps[2],/<w:keepNext\s*\/>/);assert.doesNotMatch(ps[10],/<w:keepNext\s*\/>/);assert.equal(paragraphs(xml).length,103);
});
test('uses the established Chinese and English fonts and an editable title style',()=>{
 const styles=parts(buildDocx('周会\n正文')).get('word/styles.xml');
 assert.match(styles,/w:eastAsia="KaiTi"/);assert.match(styles,/w:ascii="Times New Roman"/);assert.match(styles,/w:styleId="Title"/);
});
test('names files from the current script title and handles Windows-invalid names',()=>{
 assert.equal(wordFilename('2026年11月11日周会\n正文'),'2026年11月11日周会_司仪稿.docx');
 assert.equal(wordFilename('A/B:C*?\n正文'),'ABC_司仪稿.docx');
 assert.equal(wordFilename(' \n\t'),'周会司仪稿.docx');
});
test('paragraph properties follow the OOXML schema sequence in document and styles',()=>{
 const files=parts(buildDocx('周会\n\n在演讲比赛中，\n测试同学 荣获冠军'));
 // Ordered subset from CT_PPr: a misplaced child may trigger Word repair.
 const sequence=['pStyle','keepNext','keepLines','widowControl','spacing','jc'];
 for(const name of ['word/document.xml','word/styles.xml']){
  for(const p of files.get(name).matchAll(/<w:pPr>([\s\S]*?)<\/w:pPr>/g)){
   const children=[...p[1].matchAll(/<w:(\w+)\b/g)].map(x=>x[1]).filter(x=>sequence.includes(x));
   const indexes=children.map(x=>sequence.indexOf(x));
   assert.deepEqual(indexes,[...indexes].sort((a,b)=>a-b),`${name}: ${children.join(', ')}`);
  }
 }
});
test('keeps a repeated external competition representative intro with the next recipient',()=>{
 const xml=parts(buildDocx('周会\n\n我校机器人队代表\n高一文忠 测试同学 荣获冠军')).get('word/document.xml');
 const ps=[...xml.matchAll(/<w:p(?:\s[^>]*)?>([\s\S]*?)<\/w:p>/g)].map(x=>x[1]);
 assert.match(ps[2],/<w:keepNext\s*\/>/);
});
