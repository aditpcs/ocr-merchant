<script>
  import * as XLSX from 'xlsx';

  const headers = [
    "No.",
    "Tgl Request",
    "Peruntukkan EDC",
    "Kartu Debet BRI",
    "Kartu Kredit BRI",
    "Kartu Brizzi",
    "Kartu Bank Lain",
    "Jarkom EDC",
    "Merchant",
    "Alamat",
    "Kota",
    "Telp.",
    "MID",
    "TID",
    "Status JOB",
    "Tgl Implementasi (Max)",
    "Group",
    "Kanwil",
    "Uker Perekomendasi",
    "Kanwil Tujuan",
    "Uker Tujuan"
  ];

  let files = [];
  let previews = [];
  let processing = false;
  let status = '';
  let workbookReady = false;
  let workbook = null;

  let apiKey = localStorage.getItem('openai_api_key') || '';
  let showKeyModal = !apiKey;
  let tempKey = apiKey;

  function saveKey() {
    if (!tempKey || tempKey.trim().length < 10) return;
    apiKey = tempKey.trim();
    localStorage.setItem('openai_api_key', apiKey);
    showKeyModal = false;
  }

  function changeKey() {
    tempKey = apiKey;
    showKeyModal = true;
  }

  function handleFiles(event) {
    const selected = Array.from(event.target.files || []);
    const valid = selected.filter(f => ['image/png', 'image/jpeg', 'image/jpg'].includes(f.type));
    const merged = [...files, ...valid];
    if (merged.length > 10) {
      status = 'Maximum 10 images allowed';
      files = merged.slice(0, 10);
    } else {
      files = merged;
      status = '';
    }
    previews = files.map(file => ({
      name: file.name,
      url: URL.createObjectURL(file)
    }));
    workbookReady = false;
  }

  function clearAll() {
    files = [];
    previews = [];
    status = '';
    workbookReady = false;
    workbook = null;
  }

  function formatDateIndo(date) {
    const months = [
      'Januari','Februari','Maret','April','Mei','Juni',
      'Juli','Agustus','September','Oktober','November','Desember'
    ];
    const d = date.getDate();
    const m = months[date.getMonth()];
    const y = date.getFullYear();
    return `${d} ${m} ${y}`;
  }

  function normalizeCity(city) {
    if (!city) return '';
    let c = city.toUpperCase().replace(/\s+/g, ' ').trim();
    if (c.includes('SURAKARTA')) return 'KOTA SURAKARTA';
    c = c.replace(/KOT\.?/g, 'KOTA');
    if (c.startsWith('KOTA')) return c;
    if (c.startsWith('KAB')) return c;
    return `KOTA ${c}`;
  }

  function formatAddress({ alamat, kecamatan, kelurahan, kodepos }) {
    const parts = [alamat, kecamatan, kelurahan, kodepos].filter(Boolean).map(p => p.toUpperCase().trim());
    return parts.join(', ');
  }

  function toBase64(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result);
      reader.onerror = reject;
      reader.readAsDataURL(file);
    });
  }

  async function ocrImage(base64Image) {
    const prompt = `You are an OCR and structured data extractor. Extract the following fields from the image and return ONLY valid JSON with these exact keys: mid, tid, merchant, alamat, kecamatan, kelurahan, kodepos, kota, phone. If a field is not present, return an empty string.`;

    const body = {
      model: "gpt-4o-mini",
      temperature: 0,
      response_format: { type: "json_object" },
      messages: [
        { role: "system", content: "You extract merchant information as JSON." },
        {
          role: "user",
          content: [
            { type: "text", text: prompt },
            { type: "image_url", image_url: { url: base64Image } }
          ]
        }
      ]
    };

    const res = await fetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${apiKey}`
      },
      body: JSON.stringify(body)
    });

    if (!res.ok) {
      const err = await res.text();
      throw new Error(err);
    }
    const data = await res.json();
    const content = data.choices?.[0]?.message?.content || '{}';
    return JSON.parse(content);
  }

  function buildWorkbook(rows) {
    const row1 = Array(headers.length).fill('');
    row1[0] = 'NO SURAT SIK: SIK MANUAL';
    const row2 = Array(headers.length).fill('');
    row2[0] = 'Lampiran :';
    const sheet1Data = [row1, row2, headers, ...rows];
    const sheet2Data = [row1, row2, headers];

    const ws1 = XLSX.utils.aoa_to_sheet(sheet1Data);
    const ws2 = XLSX.utils.aoa_to_sheet(sheet2Data);

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws1, 'Lampiran SIK NON PAMERAN');
    XLSX.utils.book_append_sheet(wb, ws2, 'Lampiran SIK PAMERAN');
    return wb;
  }

  async function processImages() {
    if (!apiKey) {
      showKeyModal = true;
      return;
    }
    if (files.length === 0) return;
    processing = true;
    status = 'Processing images...';
    workbookReady = false;

    const today = new Date();
    const tglRequest = formatDateIndo(today);
    const tglMax = formatDateIndo(new Date(today.getTime() + 2 * 24 * 60 * 60 * 1000));

    const rows = [];

    try {
      for (let i = 0; i < files.length; i++) {
        status = `Processing ${i + 1} of ${files.length}...`;
        const base64 = await toBase64(files[i]);
        const ocr = await ocrImage(base64);

        const alamat = formatAddress({
          alamat: ocr.alamat,
          kecamatan: ocr.kecamatan,
          kelurahan: ocr.kelurahan,
          kodepos: ocr.kodepos
        });

        const row = [
          i + 1,
          tglRequest,
          'MERCHANT',
          'Ya',
          'Ya',
          'Ya',
          'Ya',
          'GPRS',
          (ocr.merchant || '').toUpperCase(),
          alamat,
          normalizeCity(ocr.kota),
          ocr.phone || '',
          ocr.mid || '',
          ocr.tid || '',
          'Pemasangan',
          tglMax,
          'RITEL',
          '',
          '',
          'Kas Kanpus',
          'PT. PASIFIK CIPTA SOLUSI'
        ];
        rows.push(row);
      }

      workbook = buildWorkbook(rows);
      workbookReady = true;
      status = 'Processing complete. You can download the Excel file.';
    } catch (err) {
      status = `Error: ${err.message}`;
    } finally {
      processing = false;
    }
  }

  function downloadExcel() {
    if (!workbook) return;
    XLSX.writeFile(workbook, 'merchant-ocr.xlsx');
  }
</script>

{#if showKeyModal}
  <div class="modal-overlay">
    <div class="modal">
      <h2>Enter OpenAI API Key</h2>
      <p>We need your API key to process the images.</p>
      <input type="password" placeholder="sk-..." bind:value={tempKey} />
      <button class="btn" on:click={saveKey}>Save API Key</button>
      <div class="small" style="margin-top:10px;">Key is stored in local storage.</div>
    </div>
  </div>
{/if}

<div class="container">
  <div class="header">
    <h1>Merchant OCR to Excel</h1>
    <button class="btn secondary" on:click={changeKey}>Change API Key</button>
  </div>

  <div class="card">
    <div class="upload-area">
      <input type="file" accept="image/png, image/jpeg, image/jpg" multiple on:change={handleFiles} />
      <div class="note">Maximum 10 images allowed</div>
      <div class="table-note">Accepted types: PNG, JPG, JPEG</div>
    </div>

    {#if previews.length > 0}
      <div class="preview-grid">
        {#each previews as p}
          <div class="preview-item">
            <img src={p.url} alt={p.name} />
            <div class="meta">{p.name}</div>
          </div>
        {/each}
      </div>
    {/if}

    <div class="actions">
      <button class="btn" on:click={processImages} disabled={processing || files.length === 0}>Process Images</button>
      <button class="btn secondary" on:click={clearAll} disabled={processing}>Clear</button>
      {#if workbookReady}
        <button class="btn" on:click={downloadExcel}>Download Excel</button>
      {/if}
    </div>

    {#if status}
      <div class="status">{status}</div>
    {/if}
  </div>
</div>
