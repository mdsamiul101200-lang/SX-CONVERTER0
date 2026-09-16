const express = require("express");
const cors = require("cors");
const path = require("path");
const fs = require("fs-extra");
const { v4: uuid } = require("uuid");
const multer = require("multer");
const { convert } = require("./services/conversionService");

const app = express();
const PORT = Number(process.env.PORT || 10000);
const HOST = process.env.HOST || "0.0.0.0";
const MAX_MB = Number(process.env.MAX_FILE_SIZE_MB || 100);
const DATA = path.resolve(process.env.DATA_DIR || "./data");
const uploads = path.join(DATA,"uploads"), outputs = path.join(DATA,"outputs"), tmp = path.join(DATA,"tmp");
fs.ensureDirSync(uploads); fs.ensureDirSync(outputs); fs.ensureDirSync(tmp);

app.use(cors({origin: process.env.CORS_ORIGIN || "*"}));
app.use(express.json({limit:"1mb"}));
app.use(express.static(path.join(__dirname,"..")));

const jobs = new Map();
const upload = multer({
  dest: uploads,
  limits:{fileSize:MAX_MB*1024*1024, files:1},
  fileFilter:(req,file,cb)=>{
    if(!file.originalname || file.originalname.length>180) return cb(new Error("Invalid filename."));
    cb(null,true);
  }
});

const FORMATS=["PDF","HTML","APK","ZIP","DOCX","TXT"];

app.get("/api/health",(req,res)=>res.json({ok:true,name:"SX CONVERTER",formats:FORMATS}));
app.get("/api/formats",(req,res)=>res.json({formats:FORMATS}));

app.post("/api/upload",upload.single("file"),async(req,res)=>{
  try{
    if(!req.file) return res.status(400).json({error:"No file uploaded."});
    const id=uuid();
    const safeName=path.basename(req.file.originalname).replace(/[^\w.\-() ]/g,"_").slice(0,160);
    const target=path.join(uploads,`${id}__${safeName}`);
    await fs.move(req.file.path,target,{overwrite:true});
    const detected=await fileTypeFromFile(target).catch(()=>null);
    res.json({fileId:id,originalName:safeName,size:(await fs.stat(target)).size,detectedMime:detected?.mime||null});
  }catch(e){ if(req.file?.path) await fs.remove(req.file.path).catch(()=>{}); res.status(400).json({error:e.message||"Upload failed."}); }
});

app.post("/api/convert",async(req,res)=>{
  const {fileId,inputFormat,outputFormat}=req.body||{};
  if(!fileId || !FORMATS.includes(inputFormat) || !FORMATS.includes(outputFormat))
    return res.status(400).json({error:"Invalid conversion request."});
  const matches=await fs.readdir(uploads);
  const file=matches.find(x=>x.startsWith(`${fileId}__`));
  if(!file) return res.status(404).json({error:"Uploaded file not found."});
  const jobId=uuid();
  const job={jobId,fileId,inputFormat,outputFormat,status:"QUEUED",progress:0,message:"Queued",outputPath:null,outputName:null,error:null};
  jobs.set(jobId,job);
  res.json({jobId,status:job.status});
  setImmediate(async()=>runJob(job,path.join(uploads,file)));
});

async function runJob(job,inputPath){
  job.status="PROCESSING"; job.progress=15; job.message="Inspecting input...";
  try{
    const result=await convert({job,inputPath,outputDir:outputs,tmpDir:tmp});
    job.status="COMPLETED"; job.progress=100; job.message="Conversion complete";
    job.outputPath=result.outputPath; job.outputName=result.outputName;
    const stat=await fs.stat(result.outputPath);
    if(!stat.size) throw new Error("Generated output is empty.");
  }catch(e){
    job.status="FAILED"; job.progress=0; job.error=`CONVERSION FAILED: ${e.message}`;
  }finally{
    await fs.remove(inputPath).catch(()=>{});
  }
}

app.get("/api/conversion/:id",(req,res)=>{
  const j=jobs.get(req.params.id);
  if(!j)return res.status(404).json({error:"Conversion not found."});
  res.json({jobId:j.jobId,status:j.status,progress:j.progress,message:j.message,error:j.error,outputName:j.outputName});
});

app.get("/api/download/:id",async(req,res)=>{
  const j=jobs.get(req.params.id);
  if(!j || j.status!=="COMPLETED" || !j.outputPath) return res.status(404).json({error:"Output not available."});
  if(!(await fs.pathExists(j.outputPath))) return res.status(404).json({error:"Output has expired or was removed."});
  res.download(j.outputPath,j.outputName);
});

app.delete("/api/conversion/:id",async(req,res)=>{
  const j=jobs.get(req.params.id);
  if(j?.outputPath) await fs.remove(j.outputPath).catch(()=>{});
  jobs.delete(req.params.id);
  res.json({ok:true});
});

app.use((err,req,res,next)=>{
  if(err instanceof multer.MulterError) return res.status(413).json({error:"File is too large or upload limit was exceeded."});
  res.status(400).json({error:err.message||"Request failed."});
});

app.listen(PORT, HOST, ()=>console.log(`SX CONVERTER listening on ${HOST}:${PORT}`));
