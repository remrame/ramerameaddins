// Memastikan Office JS API telah siap sebelum menjalankan kode
Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        // Menghubungkan fungsi klik ke tombol HTML
        document.getElementById("insert-text-btn").onclick = insertText;
    }
});

async function insertText() {
    try {
        await Word.run(async (context) => {
            // Mengambil elemen bodi atau area pengetikan utama di Word
            const body = context.document.body;
            
            // Memasukkan teks di akhir dokumen (End)
            body.insertText("Halo Dunia dari Office Add-in!\n", Word.InsertLocation.end);
            
            // Menjalankan perintah ke aplikasi Word
            await context.sync();
        });
    } catch (error) {
        console.error("Terjadi kesalahan: " + error);
    }
}
