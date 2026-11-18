const XLSX = require('xlsx');
const SHARE = 'https://bingol-my.sharepoint.com/:x:/g/personal/mmf_bingol_edu_tr/IQCzDZYUnSHiTrJcqePljVzNAavXXD2_cga6A3kcXmX9FJQ?e=H82dSV';
exports.handler = async () => {
  try {
    const url = new URL(process.env.EXCEL_SHARE_URL || SHARE);
    url.searchParams.set('download','1');
    const response = await fetch(url.toString(), {redirect:'follow', headers:{'User-Agent':'Mozilla/5.0'},signal:AbortSignal.timeout(22000)});
    if (!response.ok) throw new Error('OneDrive HTTP '+response.status);
    const data = Buffer.from(await response.arrayBuffer());
    if (data.length < 4 || data[0] !== 0x50 || data[1] !== 0x4b) throw new Error('OneDrive Excel yerine oturum/önizleme sayfası döndürdü');
    const book = XLSX.read(data,{type:'buffer'});
    const sheet = book.Sheets['Akademik Faaliyetler'] || book.Sheets[book.SheetNames[0]];
    const matrix = XLSX.utils.sheet_to_json(sheet,{header:1,defval:'',raw:false});
    const records = matrix.slice(1).filter(r=>String(r[5]||'').trim()).map(r=>Array.from({length:10},(_,i)=>String(r[i]??'')));
    return {statusCode:200,headers:{'Content-Type':'application/json; charset=utf-8','Cache-Control':'public, max-age=120, s-maxage=300'},body:JSON.stringify({records,source:'live',updated:new Date().toISOString()})};
  } catch (e) {
    return {statusCode:502,headers:{'Content-Type':'application/json'},body:JSON.stringify({error:String(e.message||e)})};
  }
};
