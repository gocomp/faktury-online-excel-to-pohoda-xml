<?php
if ($_SERVER["REQUEST_METHOD"] == "POST") {
    $target_dir = "uploads/";
    $target_file = $target_dir . basename($_FILES["fileToUpload"]["name"]);
    $uploadOk = 1;
    $fileType = strtolower(pathinfo($target_file,PATHINFO_EXTENSION));

    // Ellenőrizzük, hogy a fájl valóban XLSX-e
    if($fileType != "xlsx") {
        echo "Csak XLSX fájlok engedélyezettek!";
        $uploadOk = 0;
    }

    // Ellenőrizzük, hogy a fájl már létezik-e
    if (file_exists($target_file)) {
        echo "A fájl már létezik.";
        $uploadOk = 0;
    }

    // Ellenőrizzük a fájl méretét
    if ($_FILES["fileToUpload"]["size"] > 500000) {
        echo "A fájl túl nagy.";
        $uploadOk = 0;
    }

    // Engedélyezzük a feltöltést, ha minden rendben van
    if ($uploadOk == 0) {
        echo "Sikertelen feltöltés.";
    } else {
        if (move_uploaded_file($_FILES["fileToUpload"]["tmp_name"], $target_file)) {
            echo "A fájl feltöltése sikeres.";
            include 'converter.php';
            convert_to_xml($target_file);
        } else {
            echo "Hiba történt a feltöltés során.";
        }
    }
}
?>
