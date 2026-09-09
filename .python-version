from pathlib import Path
import zipfile, textwrap, json, os

out = Path("/mnt/data/gudang_web")
out.mkdir(exist_ok=True)

html = r'''<!doctype html>
<html lang="id">
<head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Sistem Gudang</title>
<style>
*{box-sizing:border-box}body{margin:0;font-family:Arial,sans-serif;background:#f4f6f8;color:#17202a}
header{background:#17324d;color:white;padding:18px 24px;font-size:22px;font-weight:700}
nav{background:white;padding:10px 20px;display:flex;gap:8px;flex-wrap:wrap;border-bottom:1px solid #ddd}
button{border:0;border-radius:8px;padding:10px 14px;cursor:pointer;background:#17324d;color:white}
nav button{background:#e9eef3;color:#17324d}.active{background:#17324d!important;color:white!important}
main{max-width:1200px;margin:24px auto;padding:0 16px}.card{background:white;border-radius:12px;padding:20px;margin-bottom:18px;box-shadow:0 2px 8px #00000010}
.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:14px}.stat{font-size:28px;font-weight:700}.muted{color:#667085}
label{font-size:13px;font-weight:700;display:block;margin-bottom:5px}input,select,textarea{width:100%;padding:10px;border:1px solid #ccd3da;border-radius:7px}textarea{min-height:80px}
table{width:100%;border-collapse:collapse}th,td{padding:10px;border-bottom:1px solid #eee;text-align:left;font-size:13px}th{background:#f7f8fa}
.hidden{display:none}.actions{display:flex;gap:8px;flex-wrap:wrap}.danger{background:#b42318}.success{background:#087443}
#preview{max-width:180px;max-height:140px;margin-top:8px;border-radius:8px;display:none}
</style>
</head>
<body>
<header>📦 LOG BOOK SPAREPART HUT PURWOKERTO</header>
<nav>
<button class="active" onclick="show('dashboard',this)">Dashboard</button>
<button onclick="show('form',this)">Input Transaksi</button>
<button onclick="show('barang',this)">Master Barang</button>
<button onclick="show('riwayat',this)">Riwayat</button>
</nav>
<main>
<section id="dashboard">
<div class="grid">
<div class="card"><div class="muted">Transaksi Masuk</div><div class="stat" id="masuk">0</div></div>
<div class="card"><div class="muted">Transaksi Keluar</div><div class="stat" id="keluar">0</div></div>
<div class="card"><div class="muted">Jenis Barang</div><div class="stat" id="jenis">0</div></div>
<div class="card"><div class="muted">Total Stok</div><div class="stat" id="stok">0</div></div>
</div>
<div class="card"><h3>Transaksi Terbaru</h3><div id="recent"></div></div>
</section>

<section id="form" class="hidden">
<div class="card"><h2>Input Barang Masuk / Keluar</h2>
<form id="trxForm">
<div class="grid">
<div><label>Tanggal</label><input id="tanggal" type="date" required></div>
<div><label>Jenis Transaksi</label><select id="tipe"><option> </option><option>MASUK</option><option>KELUAR</option></select></div>
<div><label>FME Name</label><select id="tipe"><option> </option><option>Anas</option><option>Billy</option><option>Enjang</option><option>Fredy</option><option>Dayu</option><option>Triyanto</option><option>Ibnu</option><option>Fannis</option><option>Ricko</option><option>Saryono</option><option>Zulfikar</option><option>Zaenal Farid ZTE</option><option>Fendi Huawei</option><option>Fajar</option><option>Heru S</option></select></div>
<div><label>No MDR</label><input id="dokumen" required></div>
<div><label>Sparepart Type</label><select id="tipe"><option> </option><option>VSWd</option><option>RRU A9652D</option><option>RRU 9264 M1821</option><option>RRU 9264 M7285</option><option>VBPe2q</option><option>VBPe3q</option><option>FAN BBU</option><option>VFC</option><option>VEM</option><option>VPDc</option><option>RET</option><option>Jumper Antena To RRU</option><option>IP20C 23G TL/TH</option><option>IP20C 13G TL/TH</option><option>IP20C 15G TL/TH</option><option>IP20C 7G TL/TH</option><option>IP20C 11G TL/TH</option><option>ODU RFUC 13G TL/TH</option><option>ODU RFUC 15G TL/TH</option><option>ODU RFUC 7G TL/TH</option><option>IDU Ceragon IP20G / IDU IP10G</option><option>OBU ZTE</option><option>SRU ZTE</option><option>IDU RTN Huawei</option><option>ODU Huawei 23G TL/TH</option><option>ODU Huawei15G TL/TH</option><option>IPRAN</option><option>SMR ZTE</option><option>SMR Huawei</option><option>CSU / Display Huawei</option><option>Subrack Rectifier</option><option>Battery Lithium</option><option>OLI Genset & Filter Genset</option><option>SFP 10G 1.4 KM</option><option>SFP 25G</option><option>SFP 100G</option><option>BBU 9200</option><option>Pactcore</option><option>CPRI</option><option>Filter Oli</option></select></div>
<div><label>Nama Barang</label><input id="kode" required></div>
<div><label>Serial Number</label><input id="nama" required></div>
<div><label>Jumlah</label><input id="jumlah" type="number" min="1" required></div>
<div><label>Satuan</label><input id="satuan" placeholder="pcs, box, kg"></div>
<div><label>Supplier</label><select id="tipe"><option> </option><option>SPMS ZTE</option><option>SPMS Huawei</option><option>Mitigasi</option></select></div>
<div><label>Lokasi/Rak</label><input id="lokasi"></div>
</div>
<br><label>Keterangan</label><textarea id="ket"></textarea><br>
<label>Foto Barang / Dokumen</label><input id="foto" type="file" accept="image/*"><img id="preview">
<br><br><button class="success" type="submit">Simpan Transaksi</button>
</form></div>
</section>

<section id="barang" class="hidden">
<div class="card"><h2>Master Barang</h2>
<div class="grid"><input id="mkode" placeholder="Kode"><input id="mnama" placeholder="Nama barang"><input id="msatuan" placeholder="Satuan"><input id="mlokasi" placeholder="Lokasi"></div><br>
<button onclick="addBarang()">Tambah Barang</button><br><br><div id="barangTable"></div>
</div></section>

<section id="riwayat" class="hidden">
<div class="card"><h2>Riwayat Transaksi</h2>
<div class="grid"><input id="search" oninput="render()" placeholder="Cari dokumen/kode/nama"><select id="filter" onchange="render()"><option>SEMUA</option><option>MASUK</option><option>KELUAR</option></select><button onclick="exportExcel()">Export Excel</button></div><br>
<div id="table"></div></div>
</section>
</main>
<script>
let trx=JSON.parse(localStorage.getItem('trx')||'[]'), goods=JSON.parse(localStorage.getItem('goods')||'[]'), photoData='';
const $=id=>document.getElementById(id);
$('tanggal').value=new Date().toISOString().slice(0,10);
$('foto').onchange=e=>{let f=e.target.files[0];if(f){let r=new FileReader();r.onload=()=>{$('preview').src=r.result;$('preview').style.display='block';photoData=r.result};r.readAsDataURL(f)}};
function save(){localStorage.setItem('trx',JSON.stringify(trx));localStorage.setItem('goods',JSON.stringify(goods));update()}
function show(id,b){document.querySelectorAll('main section').forEach(x=>x.classList.add('hidden'));$(id).classList.remove('hidden');document.querySelectorAll('nav button').forEach(x=>x.classList.remove('active'));b.classList.add('active');update()}
$('trxForm').onsubmit=e=>{e.preventDefault();trx.unshift({id:Date.now(),tipe:$('tipe').value,dokumen:$('dokumen').value,tanggal:$('tanggal').value,kode:$('kode').value,nama:$('nama').value,jumlah:+$('jumlah').value,satuan:$('satuan').value,pihak:$('pihak').value,lokasi:$('lokasi').value,ket:$('ket').value,foto:photoData});photoData='';$('trxForm').reset();$('tanggal').value=new Date().toISOString().slice(0,10);$('preview').style.display='none';save();alert('Transaksi berhasil disimpan')}
function addBarang(){if(!$('mkode').value||!$('mnama').value)return alert('Kode dan nama wajib diisi');goods.push({kode:$('mkode').value,nama:$('mnama').value,satuan:$('msatuan').value,lokasi:$('mlokasi').value});save();['mkode','mnama','msatuan','mlokasi'].forEach(x=>$(x).value='')}
function stock(){let s={};trx.forEach(x=>{s[x.kode]=(s[x.kode]||0)+(x.tipe==='MASUK'?x.jumlah:-x.jumlah)});return Object.values(s).reduce((a,b)=>a+b,0)}
function update(){$('masuk').textContent=trx.filter(x=>x.tipe==='MASUK').length;$('keluar').textContent=trx.filter(x=>x.tipe==='KELUAR').length;$('jenis').textContent=new Set(trx.map(x=>x.kode)).size;$('stok').textContent=stock();$('recent').innerHTML=trx.slice(0,8).map(row).join('')||'<p class="muted">Belum ada transaksi.</p>';render();renderGoods()}
function row(x){return `<div style="padding:8px 0;border-bottom:1px solid #eee"><b>${x.tipe}</b> — ${x.dokumen} — ${x.nama} (${x.jumlah} ${x.satuan||''}) — ${x.tanggal}</div>`}
function render(){let q=($('search')?.value||'').toLowerCase(),f=$('filter')?.value||'SEMUA';let a=trx.filter(x=>(f==='SEMUA'||x.tipe===f)&&JSON.stringify(x).toLowerCase().includes(q));$('table').innerHTML=`<table><tr><th>Tanggal</th><th>Jenis</th><th>Dokumen</th><th>Kode</th><th>Barang</th><th>Jumlah</th><th>Pihak</th><th>Foto</th><th>Aksi</th></tr>${a.map(x=>`<tr><td>${x.tanggal}</td><td>${x.tipe}</td><td>${x.dokumen}</td><td>${x.kode}</td><td>${x.nama}</td><td>${x.jumlah} ${x.satuan||''}</td><td>${x.pihak||''}</td><td>${x.foto?'<a target="_blank" href="'+x.foto+'">Lihat</a>':''}</td><td><button class="danger" onclick="del(${x.id})">Hapus</button></td></tr>`).join('')}</table>`}
function del(id){if(confirm('Hapus transaksi?')){trx=trx.filter(x=>x.id!==id);save()}}
function renderGoods(){$('barangTable').innerHTML=`<table><tr><th>Kode</th><th>Nama</th><th>Satuan</th><th>Lokasi</th></tr>${goods.map(x=>`<tr><td>${x.kode}</td><td>${x.nama}</td><td>${x.satuan||''}</td><td>${x.lokasi||''}</td></tr>`).join('')}</table>`}
function exportExcel(){let rows=[['Tanggal','Jenis','Nomor Dokumen','Kode Barang','Nama Barang','Jumlah','Satuan','Supplier/Tujuan','Lokasi','Keterangan','Foto']];trx.forEach(x=>rows.push([x.tanggal,x.tipe,x.dokumen,x.kode,x.nama,x.jumlah,x.satuan,x.pihak,x.lokasi,x.ket,x.foto?'Foto tersedia':'']));let csv=rows.map(r=>r.map(v=>`"${String(v??'').replaceAll('"','""')}"`).join(';')).join('\\r\\n');let blob=new Blob(["\\ufeff"+csv],{type:'text/csv;charset=utf-8'}),a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download='laporan_gudang_'+new Date().toISOString().slice(0,10)+'.csv';a.click()}
update();
</script>
</body></html>'''
