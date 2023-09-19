# veritabani_listeleme_silme_ekleme_projesi

Merhaba :)

Bu projede MS Sql Server veritabanı kullanılarak bir kitap listeleme, silme, kaydetme, word ve excel dosyasına çevirme uygulaması yer almaktadır.

1) Öncelikle veri tabanımızda "kitaplar" isminde bir  tablo oluşturalım. Benim yaptığım tablo aşağıda verilmiştir.

    ![k1](https://github.com/AhsennurUSLU/veritabani_listeleme_silme_ekleme_projesi/assets/99485329/8fcf5ef9-2af8-4494-96fe-907bde48d49c)

2) Bu işlemin ardından aşağıdaki form yapısını oluşturabilirsiniz. Ben basit bir form tasarladım dilerseniz siz daha gelişmiş bir form yapısı da oluşturabilirsiniz.

  ![f1](https://github.com/AhsennurUSLU/veritabani_listeleme_silme_ekleme_projesi/assets/99485329/b66575a5-8730-457d-9421-e632387445f9)

3) Veritabanına bağlanmak için aşağıdaki kodu kullanabilirsiniz:

   ![t3](https://github.com/AhsennurUSLU/veritabani_listeleme_silme_ekleme_projesi/assets/99485329/bc51771d-936d-4f5f-bee9-637fb24bddc1)

 Tabi öncesinde System.Data.SqlClient sınıfını import etmelisiniz.

4) Daha sonrasında veritabanına bir kaç kayıt ekleyip görüntüleme butonu için gerekli kodları yazalım (kodlar repoda mevcut) ve kodlarımızın çalışıp çalışmadığını kontrol edelim.

   ![f2](https://github.com/AhsennurUSLU/veritabani_listeleme_silme_ekleme_projesi/assets/99485329/514f2bba-ca82-40d3-a7d4-3ec3903afdae)

 Yukarda görmüş olduğunuz çıktıda görüntüle butonu çalışıyor. Ve veritabanındaki kayıtları listeledik.

5) Aynı şekilde diğer butonlar içinde silme ve kaydetme kodlarını yazalım. Kaydetme işlemi aşağıdaki gibi çalışıyor kaydet butonuna bastığınızda veritabanınızda yeni eklenen kaydı görebilirsiniz.

https://github.com/AhsennurUSLU/veritabani_listeleme_silme_ekleme_projesi/assets/99485329/f2c678d2-5f18-4c65-b275-c614f87ab198

6) Bu işlemlerin ardından excel butonu ile verilerimizi excel dosyasına çekelim. Bunun için de  Microsoft.Office.Interop.Excel kütüphanesini projenize import etmelisiniz.

   ![f4](https://github.com/AhsennurUSLU/veritabani_listeleme_silme_ekleme_projesi/assets/99485329/7653d2e5-5ee9-48a0-a036-c0de0aa37be7)

7) Aynı şekilde word dosyası butonu ile de verilerimizi word dosyasına çekelim.
   
![f5](https://github.com/AhsennurUSLU/veritabani_listeleme_silme_ekleme_projesi/assets/99485329/6aafcf7d-7e8f-4ff4-a323-6911c02f6aa5)

  Bu kodların tamamına repodan ulaşabilirsiniz. Umarım faydalı olmuştur. Teşekkür ederim :)  
