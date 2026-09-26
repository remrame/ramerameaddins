// Memastikan Office JS API telah siap sepenuhnya sebelum menjalankan kode
Office.onReady(function (info) {
    if (info.host === Office.HostType.Word) {
        // Menggunakan addEventListener standar untuk memastikan tombol merespons di cloud browser
        const myButton = document.getElementById("insert-text-btn");
        if (myButton) {
            myButton.addEventListener("click", insertText);
        }
    }
});

async function insertText() {
    try {
        await Word.run(async (context) => {
            // Mengambil area bodi dokumen utama
            const body = context.document.body;
            
            // Menggunakan Word.InsertLocation.end untuk memasukkan teks ke dokumen
            body.insertText("Halo Dunia dari Office Add-in yang Diperbarui!\n", Word.InsertLocation.end);
            
            // Menjalankan perintah dan sinkronisasi ke aplikasi Word
            await context.sync();
        });
    } catch (error) {
        console.error("Terjadi kesalahan sistem: " + error);
    }
}
