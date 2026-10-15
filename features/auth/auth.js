(function(){
 'use strict';
 const A={ready:false,user:null,client:null};let hooks,loading=false,recovery=/type=recovery/.test(location.hash),busy=false,pendingSave=null,revision=0,dirty=false,activeSave=null;
 const screen=()=>document.getElementById('bd-auth');
 function message(text,error=false){const el=document.getElementById('bd-auth-message');if(el){el.textContent=text;el.className=error?'bd-auth-error':'bd-auth-message';}}
 function panel(mode='login',text=''){
  document.body.classList.add('bd-auth-locked');screen().hidden=false;
  screen().innerHTML=`<div class="bd-auth-card"><h1>BDFlow</h1><p>${mode==='password'?'Nueva contraseña':'Acceso privado'}</p>${mode==='loading'?'<p role="status">Comprobando acceso…</p>':`<form id="bd-auth-form"><label>${mode==='password'?'Contraseña nueva':'Correo'}<input name="${mode==='password'?'password':'email'}" type="${mode==='password'?'password':'email'}" required ${mode==='password'?'minlength="12" autocomplete="new-password"':'autocomplete="username"'}></label>${mode==='login'?'<label>Contraseña<input name="password" type="password" required autocomplete="current-password"></label>':''}<button type="submit">${mode==='password'?'Guardar contraseña':mode==='reset'?'Enviar enlace de recuperación':'Entrar'}</button></form><button class="bd-auth-secondary" id="bd-auth-secondary">${mode==='login'?'He olvidado mi contraseña':'Volver'}</button>`}<p id="bd-auth-message" role="status"></p></div>`;
  message(text);document.getElementById('bd-auth-form')?.addEventListener('submit',event=>{event.preventDefault();submit(mode,event.currentTarget);});document.getElementById('bd-auth-secondary')?.addEventListener('click',()=>{if(mode==='login')panel('reset');else if(mode==='password'&&A.user&&!recovery)unlock();else panel('login');});
 }
 function wipe(){revision++;A.ready=false;A.user=null;dirty=false;clearTimeout(pendingSave);hooks?.clear();document.querySelectorAll('.modal-wrap,.rf-modal,#smartAlertPanel,.bdflow-toast').forEach(el=>{if(el.classList.contains('modal-wrap'))el.classList.remove('open','show');else el.remove();});}
 async function open(session){
  if(loading||!session?.user)return;loading=true;const ticket=++revision;A.ready=false;panel('loading');try{
   const {data,error}=await A.client.from('bdflow_access').select('workspace').eq('user_id',session.user.id).maybeSingle();if(ticket!==revision)return;if(error)throw error;if(!data)throw Error('Esta cuenta no tiene acceso a BDFlow. Utiliza tu cuenta del plan Reforma Laboral.');
   A.user=session.user;await hooks.hydrate(data.workspace);if(ticket!==revision)return;A.ready=true;if(recovery)panel('password');else unlock();
  }catch(e){wipe();panel('login',e.message||'No se pudo cargar BDFlow. Vuelve a intentarlo.');}finally{loading=false;}
 }
 function unlock(){screen().hidden=true;document.body.classList.remove('bd-auth-locked');hooks.render();account();hooks.localNotice?.();}
 function account(){let el=document.getElementById('bd-account');if(!el){el=document.createElement('div');el.id='bd-account';el.className='bd-account';document.querySelector('.main').prepend(el);}el.innerHTML='<span></span><button type="button">Cambiar contraseña</button><button type="button">Cerrar sesión</button>';el.querySelector('span').textContent=A.user?.email||'Cuenta privada';el.querySelectorAll('button')[0].onclick=()=>panel('password');el.querySelectorAll('button')[1].onclick=logout;}
 async function submit(mode,form){if(busy)return;busy=true;const btn=form.querySelector('button');btn.disabled=true;message('Procesando…');try{
  if(mode==='login'){const {data,error}=await A.client.auth.signInWithPassword({email:form.email.value.trim(),password:form.password.value});if(error)throw Error('Correo o contraseña incorrectos, o correo pendiente de confirmar.');await open(data.session);}
  else if(mode==='reset'){const {error}=await A.client.auth.resetPasswordForEmail(form.email.value.trim(),{redirectTo:location.origin+location.pathname});if(error)throw error;message('Si el correo corresponde a una cuenta, recibirás un enlace para cambiar la contraseña.');}
  else {const {error}=await A.client.auth.updateUser({password:form.password.value});if(error)throw error;recovery=false;unlock();}
 }catch(e){message(e.message||'No se pudo completar la operación.',true);}finally{busy=false;if(btn.isConnected)btn.disabled=false;}}
 async function logout(){if(busy)return;busy=true;try{await A.flush();const {error}=await A.client.auth.signOut();if(error)throw error;wipe();panel('login');}catch(e){alert('No se ha cerrado la sesión: '+e.message);}finally{busy=false;}}
 A.boot=async function(config){hooks=config;A.client=config.client;panel('loading');if(!A.client){panel('login','No se pudo cargar la conexión segura. Recarga la página.');return;}
  A.client.auth.onAuthStateChange((event,session)=>{if(event==='SIGNED_OUT'){wipe();panel('login');}else if(event==='PASSWORD_RECOVERY'){recovery=true;setTimeout(()=>open(session),0);}else if(event==='SIGNED_IN'&&session?.user.id!==A.user?.id){if(A.ready)wipe();if(!loading&&!busy)setTimeout(()=>open(session),0);}});
  try{const {data,error}=await A.client.auth.getSession();if(error)throw error;if(data.session)await open(data.session);else panel('login');}catch(e){panel('login','No se pudo comprobar la sesión. Recarga la página.');}
 };
 A.schedule=function(){if(!A.ready)return;dirty=true;clearTimeout(pendingSave);pendingSave=setTimeout(()=>A.flush().catch(e=>hooks.saveError(e)),900);};
 A.flush=async function(){clearTimeout(pendingSave);pendingSave=null;if(!A.ready)throw Error('Inicia sesión primero.');dirty=true;if(activeSave)return activeSave;activeSave=(async()=>{try{do{dirty=false;await hooks.save();}while(dirty&&A.ready);}catch(e){dirty=true;throw e;}finally{activeSave=null;}})();return activeSave;};
 window.addEventListener?.('beforeunload',event=>{if(A.ready&&(dirty||activeSave)){event.preventDefault();event.returnValue='';}});
 A.logout=logout;A.changePassword=()=>panel('password');window.BDAuth=A;
})();
