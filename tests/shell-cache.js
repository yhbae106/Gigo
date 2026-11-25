const assert=require('node:assert/strict');
const vm=require('node:vm');

module.exports=async function verifyShellCache(script,html){
  const handlers={},deleted=[],precached=[];
  const cache={addAll:async entries=>precached.push(...entries)};
  const scope='https://yhbae106.github.io/Gigo/';
  const self={registration:{scope},clients:{claim:async()=>{}},skipWaiting:async()=>{},
    addEventListener:(name,fn)=>{handlers[name]=fn}};
  const caches={keys:async()=>['gigo-shell-v1','gigo-shell-v2','work-hub-cache','sales-forecast-static-v1'],
    delete:async name=>deleted.push(name),open:async()=>cache};
  vm.runInNewContext(script,{self,caches});
  let job;
  handlers.install({waitUntil:p=>{job=p}});
  await job;
  for(const [,src] of html.matchAll(/<script[^>]+src="([^"]+)"/g)){
    if(!src.startsWith('https:'))assert.ok(precached.includes('./'+src),src+' must be available offline');
  }
  handlers.activate({waitUntil:p=>{job=p}});
  await job;
  assert.deepEqual(deleted,['gigo-shell-v1']);
  handlers.fetch({request:{method:'GET',url:'https://yhbae106.github.io/work-hub/hub.js'},respondWith:()=>assert.fail('Gigo must leave other applications to their own cache')});
};
