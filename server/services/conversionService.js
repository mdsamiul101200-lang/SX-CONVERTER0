const path=require("path");
const fs=require("fs-extra");
const {execFile}=require("child_process");
const {promisify}=require("util");
const exec=promisify(execFile);
const {PDFDocument}=require("pdf-lib");
const JSZip=require("jszip");
const {Document,Packer,Paragraph}=require("docx");

const EXT={PDF:"pdf",HTML:"html",APK:"apk",ZIP:"zip",DOCX:"docx",TXT:"txt"};

function safeName(s){return s.replace(/[^\w.\-]+/g,"_").slice(0,120)}
async function command(name,args,options={}) {
  return exec(name,args,{
    timeout:Number(process.env.JOB_TIMEOUT_MS||300000),
    maxBuffer:8*1024*1024,
    ...options
  });
}

async function convert({job,inputPath,outputDir,tmpDir}){
  const outBase=`sx_${job.jobId}`;
  const out=path.join(outputDir,`${outBase}.${EXT[job.outputFormat]}`);
  const work=path.join(tmpDir,job.jobId);
  await fs.ensureDir(work);

  const i=job.inputFormat,o=job.outputFormat;
  job.progress=22; job.message=`Preparing ${i} → ${o}...`;

  if(i===o) throw new Error("Identical input/output formats are not processed.");

  if(o==="APK"){
    if(i==="HTML") return buildApkFromHtml(inputPath,out,work,job);
    if(i==="ZIP") return buildApkFromZip(inputPath,out,work,job);
    throw new Error("APK output requires HTML or a ZIP containing meaningful web content.");
  }

  if(i==="APK"){
    if(o==="ZIP") return apkToZip(inputPath,out,work,job);
    if(o==="HTML") return apkToHtml(inputPath,out,work,job);
    throw new Error("APK can only be safely converted to recoverable web content or an archive.");
  }

  if(i==="ZIP"){
    if(o==="HTML") return zipToHtml(inputPath,out,work,job);
    throw new Error(`ZIP → ${o} requires a meaningful source representation; arbitrary archives are not renamed.`);
  }

  if(i==="HTML"){
    if(o==="TXT") return htmlToTxt(inputPath,out,job);
    if(o==="PDF") return htmlToPdf(inputPath,out,job);
    if(o==="DOCX") return htmlToDocx(inputPath,out,job);
    if(o==="ZIP") return fileToZip(inputPath,out,work,job);
  }

  if(i==="TXT"){
    if(o==="HTML") return txtToHtml(inputPath,out,job);
    if(o==="PDF") return txtToPdf(inputPath,out,job);
    if(o==="DOCX") return txtToDocx(inputPath,out,job);
    if(o==="ZIP") return fileToZip(inputPath,out,work,job);
  }

  if(i==="DOCX"){
    if(o==="TXT") return docxToTxt(inputPath,out,job);
    if(o==="HTML") return officeToHtml(inputPath,out,job);
    if(o==="PDF") return officeToPdf(inputPath,out,job);
    if(o==="ZIP") return fileToZip(inputPath,out,work,job);
  }

  if(i==="PDF"){
    if(o==="TXT") return pdfToTxt(inputPath,out,job);
    if(o==="HTML") return pdfToHtml(inputPath,out,job);
    if(o==="DOCX") return pdfToDocx(inputPath,out,job);
    if(o==="ZIP") return fileToZip(inputPath,out,work,job);
  }

  throw new Error(`No safe genuine conversion pipeline exists for ${i} → ${o}.`);
}

async function htmlToTxt(input,out,job){
  const s=await fs.readFile(input,"utf8");
  const t=s.replace(/<script[\s\S]*?<\/script>/gi,"")
    .replace(/<style[\s\S]*?<\/style>/gi,"")
    .replace(/<[^>]+>/g," ")
    .replace(/\s+/g," ").trim();
  if(!t) throw new Error("HTML contains no recoverable text.");
  await fs.writeFile(out,t); return done(out,job);
}
async function txtToHtml(input,out,job){
  const s=await fs.readFile(input,"utf8");
  const esc=s.replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");
  await fs.writeFile(out,`<!doctype html><html><head><meta charset="utf-8"><title>SX CONVERTER</title></head><body><pre>${esc}</pre></body></html>`);
  return done(out,job);
}
async function txtToDocx(input,out,job){
  const s=await fs.readFile(input,"utf8");
  const doc=new Document({sections:[{children:s.split(/\r?\n/).map(x=>new Paragraph(x))}]});
  await fs.writeFile(out,await Packer.toBuffer(doc)); return done(out,job);
}
async function htmlToDocx(input,out,job){
  const s=await fs.readFile(input,"utf8");
  const t=s.replace(/<[^>]+>/g," ").replace(/\s+/g," ").trim();
  if(!t) throw new Error("HTML contains no recoverable text.");
  const doc=new Document({sections:[{children:[new Paragraph(t)]}]});
  await fs.writeFile(out,await Packer.toBuffer(doc)); return done(out,job);
}
async function docxToTxt(input,out,job){
  const tmp=path.join(path.dirname(out),`docx_${job.jobId}`);
  await fs.ensureDir(tmp);
  await command("libreoffice",["--headless","--convert-to","txt:Text","--outdir",tmp,input]);
  const f=(await fs.readdir(tmp)).find(x=>x.toLowerCase().endsWith(".txt"));
  if(!f) throw new Error("DOCX text extraction failed.");
  await fs.copy(path.join(tmp,f),out); return done(out,job);
}
async function officeToPdf(input,out,job){
  const tmp=path.join(path.dirname(out),`pdf_${job.jobId}`);
  await fs.ensureDir(tmp);
  await command("libreoffice",["--headless","--convert-to","pdf","--outdir",tmp,input]);
  const f=(await fs.readdir(tmp)).find(x=>x.toLowerCase().endsWith(".pdf"));
  if(!f) throw new Error("Office-to-PDF conversion failed.");
  await fs.copy(path.join(tmp,f),out); return done(out,job);
}
async function officeToHtml(input,out,job){
  const tmp=path.join(path.dirname(out),`html_${job.jobId}`);
  await fs.ensureDir(tmp);
  await command("libreoffice",["--headless","--convert-to","html","--outdir",tmp,input]);
  const f=(await fs.readdir(tmp)).find(x=>x.toLowerCase().endsWith(".html"));
  if(!f) throw new Error("Office-to-HTML conversion failed.");
  await fs.copy(path.join(tmp,f),out); return done(out,job);
}
async function htmlToPdf(input,out,job){
  const tmp=path.join(path.dirname(out),`pdf_${job.jobId}`);
  await fs.ensureDir(tmp);
  await command("libreoffice",["--headless","--convert-to","pdf","--outdir",tmp,input]);
  const f=(await fs.readdir(tmp)).find(x=>x.toLowerCase().endsWith(".pdf"));
  if(!f) throw new Error("HTML-to-PDF conversion failed.");
  await fs.copy(path.join(tmp,f),out); return done(out,job);
}
async function txtToPdf(input,out,job){return officeToPdf(input,out,job)}
async function pdfToTxt(){throw new Error("PDF-to-TXT requires a dedicated PDF text-extraction engine; this route is intentionally disabled rather than producing misleading output.")}
async function pdfToHtml(){throw new Error("PDF-to-HTML layout reconstruction requires a dedicated PDF rendering/text engine; this route is intentionally disabled rather than producing misleading output.")}
async function pdfToDocx(){throw new Error("PDF-to-DOCX requires layout-aware extraction/OCR; this route is intentionally disabled rather than producing misleading output.")}

async function fileToZip(input,out,work,job){
  const zip=new JSZip(); zip.file(path.basename(input),await fs.readFile(input));
  await fs.writeFile(out,await zip.generateAsync({type:"nodebuffer",compression:"DEFLATE"}));
  return done(out,job);
}
async function apkToZip(input,out,work,job){return fileToZip(input,out,work,job)}

async function apkToHtml(input,out,work,job){
  await safeUnzip(input,work);
  const html=findHtml(work);
  if(!html) throw new Error("APK contains no recoverable HTML/web asset.");
  await fs.copy(html,out); return done(out,job);
}
async function zipToHtml(input,out,work,job){
  await safeUnzip(input,work);
  const html=findHtml(work);
  if(!html) throw new Error("ZIP contains no index.html or recoverable HTML document.");
  await fs.copy(html,out); return done(out,job);
}

async function safeUnzip(input,dir){
  const data=await fs.readFile(input);
  const zip=await JSZip.loadAsync(data,{checkCRC32:true});
  let total=0,count=0;
  for(const [name,entry] of Object.entries(zip.files)){
    count++;
    if(count>5000) throw new Error("Archive contains too many entries.");
    const normalized=path.posix.normalize(name.replace(/\\/g,"/"));
    if(normalized.startsWith("../")||normalized.startsWith("/")||normalized.includes("/../"))
      throw new Error("Unsafe archive path.");
    if(entry.dir) continue;
    const buf=await entry.async("nodebuffer"); total+=buf.length;
    if(total>300*1024*1024) throw new Error("Archive expands beyond safety limit.");
    const target=path.join(dir,normalized);
    await fs.ensureDir(path.dirname(target)); await fs.writeFile(target,buf);
  }
}
function findHtml(dir){
  let found=null;
  function walk(d){
    if(found)return;
    for(const n of fs.readdirSync(d)){
      const p=path.join(d,n),st=fs.statSync(p);
      if(st.isDirectory()) walk(p);
      else if(n.toLowerCase()==="index.html"){found=p;return;}
    }
  }
  walk(dir);
  if(found)return found;
  function walkAny(d){
    if(found)return;
    for(const n of fs.readdirSync(d)){
      const p=path.join(d,n),st=fs.statSync(p);
      if(st.isDirectory()) walkAny(p);
      else if(n.toLowerCase().endsWith(".html")){found=p;return;}
    }
  }
  walkAny(dir); return found;
}

async function buildApkFromHtml(input,out,work,job){
  const html=await fs.readFile(input,"utf8");
  if(!/<html[\s>]/i.test(html)) throw new Error("Input is not recognizable HTML.");
  const site=path.join(work,"site"); await fs.ensureDir(site);
  await fs.writeFile(path.join(site,"index.html"),html);
  return buildWebViewApk(site,out,work,job);
}

async function buildApkFromZip(input,out,work,job){
  const extracted=path.join(work,"site"); await fs.ensureDir(extracted);
  await safeUnzip(input,extracted);
  const html=findHtml(extracted);
  if(!html) throw new Error("ZIP contains no recoverable HTML.");
  return buildWebViewApk(path.dirname(html),out,work,job);
}

async function buildWebViewApk(siteDir,out,work,job){
  const project=path.join(work,"android");
  await fs.ensureDir(path.join(project,"app/src/main/assets"));
  await fs.ensureDir(path.join(project,"app/src/main/java/com/sxconverter/app"));
  await fs.ensureDir(path.join(project,"app/src/main/res/values"));

  await fs.writeFile(path.join(project,"settings.gradle"),
`pluginManagement { repositories { google(); mavenCentral(); gradlePluginPortal() } }
dependencyResolutionManagement { repositoriesMode.set(RepositoriesMode.FAIL_ON_PROJECT_REPOS); repositories { google(); mavenCentral() } }
rootProject.name="SXConverterBuild"
include(":app")`);

  await fs.writeFile(path.join(project,"build.gradle"),
`plugins { id "com.android.application" version "8.6.1" apply false }`);

  await fs.writeFile(path.join(project,"app/build.gradle"),
`plugins { id "com.android.application" }
android {
  namespace "com.sxconverter.app"
  compileSdk 35
  defaultConfig {
    applicationId "com.sxconverter.generated"
    minSdk 23
    targetSdk 35
    versionCode 1
    versionName "1.0"
  }
}`);

  await fs.writeFile(path.join(project,"app/src/main/AndroidManifest.xml"),
`<manifest xmlns:android="http://schemas.android.com/apk/res/android">
  <uses-permission android:name="android.permission.INTERNET"/>
  <application android:theme="@style/AppTheme" android:label="SX CONVERTER">
    <activity android:name=".MainActivity" android:exported="true">
      <intent-filter>
        <action android:name="android.intent.action.MAIN"/>
        <category android:name="android.intent.category.LAUNCHER"/>
      </intent-filter>
    </activity>
  </application>
</manifest>`);

  await fs.writeFile(path.join(project,"app/src/main/res/values/styles.xml"),
`<resources><style name="AppTheme" parent="android:style/Theme.Material.Light.NoActionBar"><item name="android:fontFamily">sans</item><item name="android:colorAccent">#35C8FF</item></style></resources>`);

  const java=`package com.sxconverter.app;
import android.app.Activity;
import android.os.Bundle;
import android.webkit.WebView;
public class MainActivity extends Activity {
  @Override public void onCreate(Bundle b) {
    super.onCreate(b);
    WebView w=new WebView(this);
    w.getSettings().setJavaScriptEnabled(true);
    w.getSettings().setDomStorageEnabled(true);
    w.getSettings().setAllowFileAccess(true);
    w.loadUrl("file:///android_asset/index.html");
    setContentView(w);
  }
}`;

  await fs.writeFile(path.join(project,"app/src/main/java/com/sxconverter/app/MainActivity.java"),java);

  // Copy the complete web site so relative CSS/JS/images referenced by index.html survive.
  await fs.copy(siteDir,path.join(project,"app/src/main/assets"),{
    filter: async (src)=>{
      const rel=path.relative(siteDir,src);
      return !rel.startsWith("..") && !rel.split(path.sep).some(part=>part.startsWith("."));
    }
  });

  job.progress=65; job.message="Building Android project with Gradle...";
  try {
    await command("gradle",["--no-daemon","--stacktrace","assembleDebug"],{cwd:project});
  } catch(e) {
    const detail=(e.stderr||e.stdout||"").toString().slice(-1800);
    throw new Error(`Android build failed. Check Android SDK/Gradle configuration. ${detail}`);
  }

  const built=path.join(project,"app/build/outputs/apk/debug/app-debug.apk");
  if(!fs.existsSync(built)) throw new Error("Gradle finished without producing an APK.");
  await validateApk(built);
  await fs.copy(built,out);
  return done(out,job);
}

async function validateApk(apk){
  const zip=await JSZip.loadAsync(await fs.readFile(apk),{checkCRC32:true});
  if(!zip.files["AndroidManifest.xml"]) throw new Error("Generated APK is missing AndroidManifest.xml.");
  if(!zip.files["classes.dex"]) throw new Error("Generated APK is missing classes.dex.");
}

function done(out,job){
  job.progress=95; job.message="Validating output...";
  return {outputPath:out,outputName:`SX-CONVERTER-${job.jobId}.${EXT[job.outputFormat]}`};
}
module.exports={convert};
