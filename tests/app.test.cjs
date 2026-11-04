const {test}=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const vm=require('node:vm');
const html=fs.readFileSync(require('node:path').join(__dirname,'../index.html'),'utf8');
const script=[...html.matchAll(/<script(?:\s[^>]*)?>([\s\S]*?)<\/script>/g)].map(m=>m[1]).join('\n');

// Browser boundaries only: run the production parser/generator and storage logic unchanged.
function app(storage=new Map(),answer=true){
  const elements=new Map(),listeners={},confirms=[];
  const element=id=>{
    if(!elements.has(id)){
      const classes=new Set(['hidden']);
      elements.set(id,{id,value:'',innerHTML:'',textContent:'',disabled:false,files:[],
        classList:{add:x=>classes.add(x),remove:x=>classes.delete(x),contains:x=>classes.has(x),toggle(x,v){v??=!classes.has(x);v?classes.add(x):classes.delete(x)}},
        appendChild(){},scrollIntoView(){},addEventListener(type,fn){listeners[id+':'+type]=fn},setAttribute(){},click(){},select(){}});
    }
    return elements.get(id);
  };
  let reloaded=false;
  const context={console,Date,Map,Set,Blob,URL,setTimeout,clearTimeout,
    document:{getElementById:element,querySelector:()=>element('firstTab'),querySelectorAll:()=>[],createElement:()=>element('created'+elements.size),addEventListener(type,fn){listeners['document:'+type]=fn}},
    window:{addEventListener(type,fn){listeners['window:'+type]=fn}},
    localStorage:{getItem:key=>storage.get(key)||null,setItem:(k,v)=>storage.set(k,v),removeItem:key=>storage.delete(key)},
    location:{reload(){reloaded=true}},confirm:message=>{confirms.push(message);return answer},alert(){},
    XLSX:{utils:{sheet_to_json:sheet=>sheet},SSF:{}},
  };
  vm.createContext(context);vm.runInContext(script,context);
  const run=code=>vm.runInContext(code,context);
  const setup=(nature='校内',awards=['冠军','第二名'])=>{
    context.fixture={Sheets:{'获奖资料':awards.map((award,i)=>({'部门':'教务处','校内/校外':nature,'比赛':'测试比赛','类型':'单项','班级':'高一测试','学生姓名':'测试同学'+i,'学号':'T'+i,'奖项':award}))}};
    run('parseAwardRows(fixture)');
    run("state.departments[0].presenter='测试老师';state.departments[0].role='老师'");
    element('dateInput').value='2026-11-04';element('flowInput').value='先讲话后颁奖';
  };
  return {run,element,context,storage,listeners,confirms,setup,get reloaded(){return reloaded}};
}

test('synonymous places receive identical rank',()=>{
  const a=app();
  for(const [left,right] of [['冠军','第一名'],['亚军','第二名'],['季军','第三名'],['殿军','第四名'],['金奖','一等奖'],['银奖','二等奖'],['铜奖','三等奖']]){
    assert.equal(a.run(`rankScore('${left}')`),a.run(`rankScore('${right}')`),`${left} / ${right}`);
  }
});
test('mixed rank names generate second place before champion',()=>{
  const a=app();a.setup();a.run('generateScript()');
  const output=a.element('output').value;assert.ok(output.indexOf('第二名')<output.indexOf('冠军'));
});
test('external awards retain source order',()=>{
  const a=app();a.setup('校外');a.run('generateScript()');
  const output=a.element('output').value;assert.ok(output.indexOf('冠军')<output.indexOf('第二名'));
});
test('unknown local awards retain their relative source order',()=>{
  const a=app();a.setup('校内',['特别纪念奖','评审推荐奖','第一名']);a.run('generateScript()');
  const text=a.element('output').value;assert.ok(text.indexOf('特别纪念奖')<text.indexOf('评审推荐奖'));
});
test('speech-only workbook accepts an absent awards worksheet',()=>{
  const a=app();a.run("state.meta.flow='纯讲话';state.speakers=[{full:'测试校长',thanks:'校长',action:'讲话'}];parseAwardRows({Sheets:{}})");
  a.element('flowInput').value='纯讲话';a.run('generateScript()');
  assert.match(a.element('output').value,/测试校长/);assert.doesNotMatch(a.element('output').value,/颁奖仪式/);
});
test('speech-only mode does not warn about unused incomplete awards',()=>{
  const a=app();a.setup('未知');a.element('flowInput').value='纯讲话';
  a.run("state.departments[0].presenter='';state.speakers=[{full:'测试老师',thanks:'老师',action:'讲话'}];generateScript()");
  assert.doesNotMatch(a.element('tab-issues').innerHTML,/没有正确填写|没有设置颁奖人/);
});
test('spaces in local/external field do not discard a valid award',()=>{
  const a=app();a.setup('校 内',['冠军']);a.run('generateScript()');assert.match(a.element('output').value,/冠军/);
});
test('invalid nature is prominently reported alongside the script',()=>{
  const a=app();a.setup('未知');a.run('generateScript()');
  const banner=a.element('generationStatus');assert.equal(banner.classList.contains('hidden'),false);assert.match(banner.innerHTML,/测试比赛|未生成/);
});
test('missing presenter is prominently reported alongside the script',()=>{
  const a=app();a.setup();a.run("state.departments[0].presenter='';generateScript()");
  assert.match(a.element('generationStatus').innerHTML,/颁奖人/);
});
test('manual edits are preserved when replacement is cancelled',()=>{
  const a=app(new Map(),false);a.setup();a.run('generateScript()');a.element('output').value='人工修订稿';
  a.run('generateScript()');assert.equal(a.element('output').value,'人工修订稿');assert.equal(a.confirms.length,1);
});
test('confirmed replacement generates a new draft',()=>{
  const a=app();a.setup();a.run('generateScript()');a.element('output').value='人工修订稿';a.run('generateScript()');
  assert.match(a.element('output').value,/冠军/);assert.equal(a.confirms.length,1);
});
test('unchanged generated text can be regenerated without confirmation',()=>{
  const a=app();a.setup();a.run('generateScript();generateScript()');assert.equal(a.confirms.length,0);
});
test('local draft restores source settings and manual output after refresh',()=>{
  const a=app();a.setup();a.run('generateScript()');a.element('output').value='人工修订稿';a.run('saveDraft()');
  const b=app(a.storage);assert.equal(b.element('output').value,'人工修订稿');assert.equal(b.run('state.events.length'),1);
  assert.equal(b.element('outputCard').classList.contains('hidden'),false);
});
test('restored draft still asks before overwriting a manual edit',()=>{
  const a=app();a.setup();a.run('generateScript()');a.element('output').value='人工修订稿';a.run('saveDraft()');
  const b=app(a.storage,false);b.run('generateScript()');assert.equal(b.element('output').value,'人工修订稿');assert.equal(b.confirms.length,1);
});
test('cancelled reset preserves the current draft',()=>{
  const a=app(new Map(),false);a.setup();a.run('generateScript();saveDraft();resetAll()');assert.equal(a.reloaded,false);assert.ok(a.storage.size>0);
});
test('confirmed reset removes the draft',()=>{
  const a=app();a.setup();a.run('generateScript();saveDraft();resetAll()');assert.equal(a.reloaded,true);assert.equal(a.storage.size,0);
});
test('storage errors do not prevent generating or editing',()=>{
  const a=app();a.context.localStorage.setItem=()=>{throw Error('quota')};a.setup();a.run('generateScript();saveDraft()');
  assert.match(a.element('output').value,/冠军/);assert.match(a.element('draftStatus').textContent,/无法|未能|不能/);
});
test('before-unload warns if edited work cannot be saved',()=>{
  const a=app();a.context.localStorage.setItem=()=>{throw Error('quota')};a.setup();a.run('generateScript()');
  let prevented=false;const event={preventDefault(){prevented=true}};a.listeners['window:beforeunload'](event);assert.equal(prevented,true);
});

test('cancelled import preserves the existing draft and source data',()=>{
  const a=app(new Map(),false);a.setup();a.run('generateScript()');a.element('output').value='人工修订稿';
  a.run("readFile({name:'new.xlsx'})");assert.equal(a.element('output').value,'人工修订稿');assert.equal(a.run('state.events.length'),1);
});
test('failed import preserves the current source data and saved manual draft',()=>{
  const a=app();a.setup();a.run('generateScript()');a.element('output').value='人工修订稿';a.run('saveDraft()');
  a.context.FileReader=class{readAsArrayBuffer(){this.onload({target:{result:new ArrayBuffer(0)}})}};
  a.context.XLSX.read=()=>{throw Error('损坏文件')};a.run("readFile({name:'broken.xlsx'})");
  assert.equal(a.element('output').value,'人工修订稿');assert.equal(a.run('state.events.length'),1);
  assert.match(a.element('readIssues').innerHTML,/损坏文件/);assert.equal(app(a.storage).element('output').value,'人工修订稿');
});
test('successful import discards old output and persists only the newly loaded workbook',()=>{
  const a=app();a.setup();a.run('generateScript()');a.element('output').value='旧人工稿';
  a.context.FileReader=class{readAsArrayBuffer(){this.onload({target:{result:new ArrayBuffer(0)}})}};
  a.context.XLSX.read=()=>({Sheets:{'周会设置':[['流程','纯讲话'],['顺序','上台完整称呼','感谢时称呼'],[1,'新讲话人','老师']]}});
  a.run("readFile({name:'speech.xlsx'})");assert.equal(a.element('output').value,'');assert.equal(a.run('state.events.length'),0);
  const restored=app(a.storage);assert.equal(restored.element('output').value,'');assert.equal(restored.run('state.speakers[0].full'),'新讲话人');
});
test('unrelated workbook is rejected without erasing current work',()=>{
  const a=app();a.setup();a.run('generateScript()');const prior=a.element('output').value;
  a.context.FileReader=class{readAsArrayBuffer(){this.onload({target:{result:new ArrayBuffer(0)}})}};
  a.context.XLSX.read=()=>({Sheets:{'Sheet1':[['random']]}});a.run("readFile({name:'unrelated.xlsx'})");
  assert.equal(a.element('output').value,prior);assert.match(a.element('readIssues').innerHTML,/标准模板/);
});
test('an empty awards sheet without settings cannot erase the current draft',()=>{
  const a=app();a.setup();a.run('generateScript()');const prior=a.element('output').value;
  a.context.FileReader=class{readAsArrayBuffer(){this.onload({target:{result:new ArrayBuffer(0)}})}};
  a.context.XLSX.read=()=>({Sheets:{'获奖资料':[]}});a.run("readFile({name:'empty.xlsx'})");
  assert.equal(a.element('output').value,prior);assert.equal(a.run('state.events.length'),1);assert.match(a.element('readIssues').innerHTML,/读取失败/);
});
test('unrecognized award column names cannot erase the current draft',()=>{
  const a=app();a.setup();a.run('generateScript()');const prior=a.element('output').value;
  a.context.FileReader=class{readAsArrayBuffer(){this.onload({target:{result:new ArrayBuffer(0)}})}};
  a.context.XLSX.read=()=>({Sheets:{'获奖资料':[{'random':'data'}]}});a.run("readFile({name:'wrong-columns.xlsx'})");
  assert.equal(a.element('output').value,prior);assert.equal(a.run('state.events.length'),1);
});
test('new manual edits made while reading a workbook survive when replacement is cancelled',()=>{
  const a=app();a.setup();a.run('generateScript()');let reader;
  a.context.FileReader=class{readAsArrayBuffer(){reader=this}};
  a.context.XLSX.read=()=>({Sheets:{'周会设置':[['流程','纯讲话']]}});
  a.run("readFile({name:'slow.xlsx'})");a.element('output').value='读取期间的新修改';
  a.context.confirm=()=>false;reader.onload({target:{result:new ArrayBuffer(0)}});
  assert.equal(a.element('output').value,'读取期间的新修改');assert.equal(a.run('state.events.length'),1);
});
test('corrupt stored JSON does not prevent a fresh generation',()=>{
  const a=app(new Map([['mc-draft:v1','not-json']]));a.setup();a.run('generateScript()');assert.match(a.element('output').value,/冠军/);
});
test('editing the output saves it immediately through the page input event',()=>{
  const a=app();a.setup();a.run('generateScript()');a.element('output').value='已修改';
  a.listeners['document:input']({target:{id:'output',type:'textarea'}});assert.equal(app(a.storage).element('output').value,'已修改');
});
test('last row alone controls no-photo behavior for merged competitions',()=>{
  const a=app();a.setup();a.context.fixture.Sheets['获奖资料'][0]['不合影']='是';
  a.run('parseAwardRows(fixture)');assert.equal(a.run('state.events[0].noPhoto'),false);
  a.context.fixture.Sheets['获奖资料'][1]['不合影']='是';a.run('parseAwardRows(fixture)');assert.equal(a.run('state.events[0].noPhoto'),true);
});
test('empty awards mode exposes an error instead of a false completed award ceremony',()=>{
  const a=app();a.element('flowInput').value='纯颁奖';a.run('generateScript()');assert.match(a.element('generationStatus').innerHTML,/没有获奖资料/);
});
test('pure speech without a speaker exposes the missing input',()=>{
  const a=app();a.element('flowInput').value='纯讲话';a.run('generateScript()');assert.match(a.element('generationStatus').innerHTML,/讲话人/);
});
test('an award row without a recipient cannot appear as a clean generation',()=>{
  const a=app();a.setup('校内',['冠军']);delete a.context.fixture.Sheets['获奖资料'][0]['班级'];delete a.context.fixture.Sheets['获奖资料'][0]['学生姓名'];
  a.run('parseAwardRows(fixture);generateScript()');assert.match(a.element('generationStatus').innerHTML,/领奖资料缺少/);
});
test('local class awards are announced from lowest to highest',()=>{
  const a=app();a.setup('校内',['第一名','第三名','第二名']);
  a.context.fixture.Sheets['获奖资料'].forEach(r=>r['类型']='班级');a.run('parseAwardRows(fixture);generateScript()');
  const text=a.element('output').value;assert.ok(text.indexOf('第三名')<text.indexOf('第二名'));assert.ok(text.indexOf('第二名')<text.indexOf('第一名'));
});
test('unknown local team awards emit a reminder',()=>{
  const a=app();a.setup('校内',['']);Object.assign(a.context.fixture.Sheets['获奖资料'][0],{'类型':'团体','团体奖项':'评审特别推荐奖','社团/校队/单位':'测试队'});
  a.run('parseAwardRows(fixture);generateScript()');assert.match(a.element('tab-issues').innerHTML,/无法判断/);
});

test('Word download exports the current manually edited text without regenerating',async()=>{
  const a=app();a.setup();a.run('generateScript()');
  a.context.MCWordExport=require('../word-export.js');
  const source='2026年11月11日周会\n\n请老师留步。手动改稿 A & B <确认>。';
  a.element('output').value=source;a.element('dateInput').value='2026-12-01';
  let blob,anchor;
  a.context.URL={createObjectURL(value){blob=value;return 'blob:test'},revokeObjectURL(){}};
  a.context.document.createElement=()=>({click(){anchor=this}});
  a.run('downloadWord()');
  assert.equal(a.element('output').value,source);
  assert.equal(anchor.download,'2026年11月11日周会_司仪稿.docx');
  assert.equal(blob.type,'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
  const bytes=Buffer.from(await blob.arrayBuffer());assert.equal(bytes.readUInt32LE(0),0x04034b50);
  assert.ok(bytes.includes(Buffer.from('请老师留步。手动改稿 A &amp; B &lt;确认&gt;。')));
});
test('Word download explains empty input without starting a download',()=>{
  const a=app(),messages=[];a.context.MCWordExport=require('../word-export.js');a.context.alert=x=>messages.push(x);
  a.element('output').value=' \n\t';a.run('downloadWord()');assert.equal(messages.length,1);assert.match(messages[0],/生成或输入司仪稿/);
});
test('Word module load failure leaves the draft intact and explains recovery',()=>{
  const a=app(),messages=[];a.element('output').value='手动稿';a.context.alert=x=>messages.push(x);
  a.run('downloadWord()');assert.equal(a.element('output').value,'手动稿');assert.match(messages[0],/刷新/);
});

function classGroupWorkbook(a,rows){
 a.context.fixture={Sheets:{'获奖资料':rows.map((r,i)=>({'学号':'G'+i,'班级':r[0],'学生姓名':'测试代表'+i,'比赛':r[1],'组别':r[2],'奖项/个人奖项':r[3],'不合影':r[4],...(i===0?{'部门':'体育处','校内/校外':'校内','类型':'班级'}:{})}))}};
 a.run("parseAwardRows(fixture);state.departments[0].presenter='体育处测试老师';state.departments[0].role='主任'");
 a.element('flowInput').value='纯颁奖';
}
test('class groups stay separate and retain their Excel order before rank sorting',()=>{
 const a=app();classGroupWorkbook(a,[
  ['初中冠军班','足球赛','初中男子组','冠军','是'],['初中殿军班','足球赛','','殿军','是'],
  ['高中冠军班','足球赛','高中男子组','冠军','是'],['高中殿军班','足球赛','','殿军',''],
 ]);a.run('generateScript()');const text=a.element('output').value;
 assert.equal(a.run('state.events.length'),2);
 for(const group of ['初中男子组','高中男子组'])assert.match(text,new RegExp(group));
 assert.ok(text.indexOf('初中殿军班')<text.indexOf('初中冠军班'));
 assert.ok(text.indexOf('初中冠军班')<text.indexOf('高中男子组'));
 assert.ok(text.indexOf('高中殿军班')<text.indexOf('高中冠军班'));
});
test('class group changes work with a blank competition cell and repeated group labels do not split',()=>{
 const a=app();classGroupWorkbook(a,[
  ['甲班','足球赛','初中组','殿军','是'],['乙班','','初中组','冠军','是'],
  ['丙班','','高中组','殿军','是'],['丁班','','','冠军',''],
 ]);
 assert.equal(a.run('state.events.length'),2);assert.equal(a.run('state.events[1].competition'),'足球赛');
 assert.equal(a.run('state.events[1].rows[1].group'),'高中组');
});
test('photo endings follow the last source row of each class group before sorting',()=>{
 const a=app();classGroupWorkbook(a,[
  ['甲冠军班','排球赛','男子组','冠军','是'],['甲殿军班','排球赛','','殿军',''],
  ['乙冠军班','排球赛','女子组','冠军',''],['乙殿军班','排球赛','','殿军','是'],
 ]);a.run('generateScript()');const text=a.element('output').value;
 assert.equal((text.match(/一同合影留念/g)||[]).length,1);
 assert.ok(text.indexOf('一同合影留念')<text.indexOf('女子组'));
});
test('class photo wording does not invent accompanying teachers',()=>{
 const a=app();classGroupWorkbook(a,[['甲班','整洁比赛','初中组','冠军','']]);a.run('generateScript()');
 const text=a.element('output').value;assert.match(text,/请获奖同学与主任一同合影留念/);assert.doesNotMatch(text,/获奖同学及老师|谢谢老师/);
});
test('a class teacher is still included in class photo wording',()=>{
 const a=app();classGroupWorkbook(a,[['甲班','整洁比赛','初中组','冠军','']]);a.context.fixture.Sheets['获奖资料'][0]['班导师']='测试班导师';
 a.run('parseAwardRows(fixture);generateScript()');assert.match(a.element('output').value,/获奖同学及班导师与主任一同合影留念/);
});
test('ungrouped class competitions remain one block',()=>{
 const a=app();classGroupWorkbook(a,[['甲班','整洁比赛','','冠军','是'],['乙班','整洁比赛','','殿军','']]);
 a.run('generateScript()');assert.equal(a.run('state.events.length'),1);assert.match(a.element('output').value,/在整洁比赛中，获奖班级有/);
});
test('regenerating an old mixed-group class block requests re-upload rather than mixing groups',()=>{
 const a=app();a.setup();a.run("state.events[0].type='班级';state.events[0].rows[0].group='初中组';state.events[0].rows[1].group='高中组';generateScript()");
 assert.match(a.element('generationStatus').innerHTML,/重新上传/);
});
test('restoring a legacy mixed class draft warns immediately and preserves manual output',()=>{
 const a=app();a.setup();a.run('generateScript()');
 a.run("state.events[0].type='班级';state.events[0].rows[0].group='初中组';state.events[0].rows[1].group='高中组';lastIssues=[]");
 a.element('output').value='旧版稿件，包含手动修改';a.run('saveDraft()');
 const b=app(a.storage);assert.equal(b.element('output').value,'旧版稿件，包含手动修改');
 assert.match(b.element('generationStatus').innerHTML,/重新上传/);assert.doesNotMatch(b.element('generationStatus').innerHTML,/没有发现明显异常/);
 b.run('saveDraft()');const c=app(b.storage);assert.equal((c.element('generationStatus').innerHTML.match(/重新上传/g)||[]).length,1);
});
test('external class groups preserve each group source order instead of ranking awards',()=>{
 const a=app();classGroupWorkbook(a,[['甲冠军班','外赛','初中组','冠军','是'],['甲殿军班','外赛','','殿军','是'],['乙冠军班','外赛','高中组','冠军','是'],['乙殿军班','外赛','','殿军','']]);
 a.context.fixture.Sheets['获奖资料'][0]['校内/校外']='校外';a.run('parseAwardRows(fixture);generateScript()');
 const text=a.element('output').value;
 assert.ok(text.indexOf('甲冠军班')<text.indexOf('甲殿军班'));assert.ok(text.indexOf('甲殿军班')<text.indexOf('高中组'));assert.ok(text.indexOf('乙冠军班')<text.indexOf('乙殿军班'));
});

module.exports={app};
