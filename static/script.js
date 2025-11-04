// Menunggu seluruh konten HTML dimuat sebelum menjalankan script
document.addEventListener("DOMContentLoaded", () => {

  const globalErrorDiv = document.getElementById("global-error");

  const proofreadFileInput = document.getElementById("proofread-file");
  const proofreadAnalyzeBtn = document.getElementById("proofread-analyze-btn");
  const proofreadLoading = document.getElementById("proofread-loading");
  const proofreadResultsContainer = document.getElementById("proofread-results-container");
  const proofreadResultsTableDiv = document.getElementById("proofread-results-table");
  const proofreadDownloadRevisedBtn = document.getElementById("proofread-download-revised-btn");
  const proofreadDownloadHighlightedBtn = document.getElementById("proofread-download-highlighted-btn");
  const proofreadDownloadZipBtn = document.getElementById("proofread-download-zip-btn");

  const compareFileInput1 = document.getElementById("compare-file1");
  const compareFileInput2 = document.getElementById("compare-file2");
  const compareAnalyzeBtn = document.getElementById("compare-analyze-btn");
  const compareLoading = document.getElementById("compare-loading");
  const compareResultsContainer = document.getElementById("compare-results-container");
  const compareResultsTableDiv = document.getElementById("compare-results-table");
  const compareDownloadBtn = document.getElementById("compare-download-btn");

  const coherenceFileInput = document.getElementById("coherence-file");
  const coherenceAnalyzeBtn = document.getElementById("coherence-analyze-btn");
  const coherenceLoading = document.getElementById("coherence-loading");
  const coherenceResultsContainer = document.getElementById("coherence-results-container");
  const coherenceResultsTableDiv = document.getElementById("coherence-results-table");
  
  const restructureFileInput = document.getElementById("restructure-file");
  const restructureAnalyzeBtn = document.getElementById("restructure-analyze-btn");
  const restructureLoading = document.getElementById("restructure-loading");
  const restructureResultsContainer = document.getElementById("restructure-results-container");
  const restructureResultsTableDiv = document.getElementById("restructure-results-table");
  const restructureDownloadBtn = document.getElementById("restructure-download-btn");

  // =============================================
  // ---         FUNGSI-FUNGSI HELPER          ---
  // =============================================

  /** Menampilkan pesan error global */
  function showError(message) {
    globalErrorDiv.textContent = `Terjadi Kesalahan: ${message}`;
    globalErrorDiv.classList.remove("hidden");
  }

  /** Menghilangkan pesan error global */
  function clearError() {
    globalErrorDiv.textContent = "";
    globalErrorDiv.classList.add("hidden");
  }

  /**
   * Membuat tabel HTML dari data JSON.
   * @param {Array<Object>} data - Array berisi objek data.
   * @param {Array<string>} headers - Array berisi nama header (key dari objek).
   * @returns {string} - String HTML untuk tabel.
   */
  function createTable(data, headers) {
    if (!data || data.length === 0) {
      return "<p>Tidak ada data untuk ditampilkan.</p>";
    }

    let a = 1
    let head = "<tr>";
    head += `<th>No.</th>`
    let customHeaders = {
        "Kata/Frasa Salah": "Salah",
        "Perbaikan Sesuai KBBI": "Perbaikan",
        "Pada Kalimat": "Konteks Kalimat",
        "Ditemukan di Halaman": "Halaman",
        "Kalimat Awal": "Kalimat Asli",
        "Kalimat Revisi": "Kalimat Revisi",
        "Kata yang Direvisi": "Perubahan",
        "topik": "Topik Utama",
        "asli": "Teks Asli",
        "saran": "Saran Revisi",
        "Paragraf yang Perlu Dipindah": "Paragraf",
        "Lokasi Asli": "Lokasi Asli",
        "Saran Lokasi Baru": "Saran Lokasi"
    };

    headers.forEach(header => {
      head += `<th>${customHeaders[header] || header}</th>`;
    });
    head += "</tr>";

    let body = "";
    data.forEach(row => {
      body += "<tr>";
      body += `<td>${a++}</td>`
      headers.forEach(header => {
        body += `<td>${row[header] || ""}</td>`;
      });
      body += "</tr>";
    });

    return `
      <div class="results-table-wrapper">
        <table>
          <thead>${head}</thead>
          <tbody>${body}</tbody>
        </table>
      </div>
    `;
  }

  /**
   * Menangani proses download file dari API.
   * @param {string} url - Endpoint API untuk download.
   * @param {FormData} formData - Data (file) yang akan dikirim.
   ** @param {string} defaultFilename - Nama file jika header tidak ada.
   */
  async function handleDownload(url, formData, defaultFilename = "download.dat") {
    clearError();
    try {
      const response = await fetch(url, {
        method: "POST",
        body: formData,
      });

      if (!response.ok) {
        const err = await response.json();
        throw new Error(err.error || "Gagal mengunduh file");
      }

      const blob = await response.blob();
      
      // Coba dapatkan nama file dari header
      const contentDisposition = response.headers.get('content-disposition');
      let filename = defaultFilename;
      if (contentDisposition) {
          const filenameRegex = /filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/;
          const matches = filenameRegex.exec(contentDisposition);
          if (matches != null && matches[1]) {
            filename = matches[1].replace(/['"]/g, '');
          }
      }

      // Buat link sementara untuk memicu download
      const link = document.createElement("a");
      link.href = window.URL.createObjectURL(blob);
      link.download = filename;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      window.URL.revokeObjectURL(link.href);

    } catch (error) {
      showError(error.message);
    }
  }

  // =============================================
  // ---      EVENT LISTENERS (FITUR 1)        ---
  // =============================================
  if (proofreadAnalyzeBtn) {
    proofreadAnalyzeBtn.addEventListener("click", async () => {
      const file = proofreadFileInput.files[0];
      if (!file) {
        showError("Silakan pilih file terlebih dahulu.");
        return;
      }
      clearError();
      proofreadLoading.classList.remove("hidden");
      proofreadAnalyzeBtn.disabled = true;
      proofreadResultsContainer.classList.add("hidden");

      const formData = new FormData();
      formData.append("file", file);

      try {
        const response = await fetch("/api/proofread/analyze", {
          method: "POST",
          body: formData,
        });

        if (!response.ok) {
          const err = await response.json();
          throw new Error(err.error || "Respon server tidak valid");
        }

        const data = await response.json();

        if (data.length === 0) {
          proofreadResultsTableDiv.innerHTML = "<p>Tidak ada kesalahan yang ditemukan.</p>";
        } else {
          const headers = ["Kata/Frasa Salah", "Perbaikan Sesuai KBBI", "Pada Kalimat", "Ditemukan di Halaman"];
          proofreadResultsTableDiv.innerHTML = createTable(data, headers);
        }
        proofreadResultsContainer.classList.remove("hidden");

      } catch (error) {
        showError(error.message);
      } finally {
        proofreadLoading.classList.add("hidden");
        proofreadAnalyzeBtn.disabled = false;
      }
    });

    // Listeners untuk tombol download Proofreading
    proofreadDownloadRevisedBtn.addEventListener("click", () => {
      const file = proofreadFileInput.files[0];
      if (!file) { showError("File asli tidak ditemukan."); return; }
      const formData = new FormData();
      formData.append("file", file);
      handleDownload("/api/proofread/download/revised", formData, `revisi_${file.name}`);
    });

    proofreadDownloadHighlightedBtn.addEventListener("click", () => {
      const file = proofreadFileInput.files[0];
      if (!file) { showError("File asli tidak ditemukan."); return; }
      const formData = new FormData();
      formData.append("file", file);
      handleDownload("/api/proofread/download/highlighted", formData, `highlight_${file.name}`);
    });

    proofreadDownloadZipBtn.addEventListener("click", () => {
      const file = proofreadFileInput.files[0];
      if (!file) { showError("File asli tidak ditemukan."); return; }
      const formData = new FormData();
      formData.append("file", file);
      handleDownload("/api/proofread/download/zip", formData, `hasil_proofread_${file.name}.zip`);
    });
  }
  
  // =============================================
  // ---      EVENT LISTENERS (FITUR 2)        ---
  // =============================================
  if (compareAnalyzeBtn) {
    compareAnalyzeBtn.addEventListener("click", async () => {
      const file1 = compareFileInput1.files[0];
      const file2 = compareFileInput2.files[0];

      if (!file1 || !file2) {
        showError("Silakan unggah KEDUA file untuk perbandingan.");
        return;
      }
      clearError();
      compareLoading.classList.remove("hidden");
      compareAnalyzeBtn.disabled = true;
      compareResultsContainer.classList.add("hidden");

      const formData = new FormData();
      formData.append("file1", file1);
      formData.append("file2", file2);

      try {
        const response = await fetch("/api/compare/analyze", {
          method: "POST",
          body: formData,
        });

        if (!response.ok) {
          const err = await response.json();
          throw new Error(err.error || "Respon server tidak valid");
        }

        const data = await response.json();

        if (data.length === 0) {
          compareResultsTableDiv.innerHTML = "<p>Tidak ada perbedaan signifikan yang ditemukan.</p>";
        } else {
          const headers = ["Kalimat Awal", "Kalimat Revisi", "Kata yang Direvisi"];
          compareResultsTableDiv.innerHTML = createTable(data, headers);
        }
        compareResultsContainer.classList.remove("hidden");

      } catch (error) {
        showError(error.message);
      } finally {
        compareLoading.classList.add("hidden");
        compareAnalyzeBtn.disabled = false;
      }
    });

    compareDownloadBtn.addEventListener("click", () => {
      const file1 = compareFileInput1.files[0];
      const file2 = compareFileInput2.files[0];
      if (!file1 || !file2) { showError("File asli tidak ditemukan."); return; }
      const formData = new FormData();
      formData.append("file1", file1);
      formData.append("file2", file2);
      handleDownload("/api/compare/download", formData, `perbandingan_${file1.name}`);
    });
  }

  // =============================================
  // ---      EVENT LISTENERS (FITUR 3)        ---
  // =============================================
  if (coherenceAnalyzeBtn) {
    coherenceAnalyzeBtn.addEventListener("click", async () => {
      const file = coherenceFileInput.files[0];
      if (!file) {
        showError("Silakan pilih file terlebih dahulu.");
        return;
      }
      clearError();
      coherenceLoading.classList.remove("hidden");
      coherenceAnalyzeBtn.disabled = true;
      coherenceResultsContainer.classList.add("hidden");

      const formData = new FormData();
      formData.append("file", file);

      try {
        const response = await fetch("/api/coherence/analyze", {
          method: "POST",
          body: formData,
        });

        if (!response.ok) {
          const err = await response.json();
          throw new Error(err.error || "Respon server tidak valid");
        }

        const data = await response.json();

        if (data.length === 0) {
          coherenceResultsTableDiv.innerHTML = "<p>Tidak ada masalah koherensi yang ditemukan.</p>";
        } else {
          const headers = ["topik", "asli", "saran"];
          coherenceResultsTableDiv.innerHTML = createTable(data, headers);
        }
        coherenceResultsContainer.classList.remove("hidden");

      } catch (error) {
        showError(error.message);
      } finally {
        coherenceLoading.classList.add("hidden");
        coherenceAnalyzeBtn.disabled = false;
      }
    });
  }
  
  // =============================================
  // ---      EVENT LISTENERS (FITUR 4)        ---
  // =============================================
  if (restructureAnalyzeBtn) {
    restructureAnalyzeBtn.addEventListener("click", async () => {
      const file = restructureFileInput.files[0];
      if (!file) {
        showError("Silakan pilih file terlebih dahulu.");
        return;
      }
      clearError();
      restructureLoading.classList.remove("hidden");
      restructureAnalyzeBtn.disabled = true;
      restructureResultsContainer.classList.add("hidden");

      const formData = new FormData();
      formData.append("file", file);

      try {
        const response = await fetch("/api/restructure/analyze", {
          method: "POST",
          body: formData,
        });

        if (!response.ok) {
          const err = await response.json();
          throw new Error(err.error || "Respon server tidak valid");
        }

        const data = await response.json();

        if (data.length === 0) {
          restructureResultsTableDiv.innerHTML = "<p>Tidak ada saran restrukturisasi.</p>";
        } else {
          const headers = ["Paragraf yang Perlu Dipindah", "Lokasi Asli", "Saran Lokasi Baru"];
          restructureResultsTableDiv.innerHTML = createTable(data, headers);
        }
        restructureResultsContainer.classList.remove("hidden");

      } catch (error) {
        showError(error.message);
      } finally {
        restructureLoading.classList.add("hidden");
        restructureAnalyzeBtn.disabled = false;
      }
    });

    restructureDownloadBtn.addEventListener("click", () => {
      const file = restructureFileInput.files[0];
      if (!file) { showError("File asli tidak ditemukan."); return; }
      const formData = new FormData();
      formData.append("file", file);
      handleDownload("/api/restructure/download", formData, `highlight_rekomendasi_${file.name}`);
    });
  }

}); 