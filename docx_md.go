package main

import (
	"archive/zip"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strings"

	"github.com/beevik/etree"
)

func main() {
	zipFilePath := "/Users/jonas/hello/test.docx"
	fileName := filepath.Base(zipFilePath)
	extension := filepath.Ext(fileName)
	fileNameWithoutExtension := fileName[:len(fileName)-len(extension)]
	directoryPath := filepath.Dir(zipFilePath)
	unzip(zipFilePath, directoryPath+"/"+fileNameWithoutExtension)

	// do stuff here
	addParagraph(directoryPath + "/" + fileNameWithoutExtension + "/word/document.xml")

	zipFolder(directoryPath+"/"+fileNameWithoutExtension, directoryPath+"/"+fileName)
	deleteFolder(directoryPath + "/" + fileNameWithoutExtension)

}

// XML

func addParagraph(docxPath string) {
	// XML-Datei laden
	doc := etree.NewDocument()
	if err := doc.ReadFromFile(docxPath); err != nil {
		fmt.Println("Error reading the XML file:", err)
		return
	}

	// Neues XML-Element erstellen
	pElement := etree.NewElement("w:p")
	rElement := etree.NewElement("w:r")
	tElement := etree.NewElement("w:t")
	pElement.AddChild(rElement)
	rElement.AddChild(tElement)
	tElement.SetText("A script has created me.")

	// Den <w:body>-Knoten finden
	body := doc.FindElement("/w:document/w:body")
	if body == nil {
		fmt.Println("Error: The <w:body> node was not found.")
		return
	}

	// Das neue XML-Element zum <w:body>-Knoten hinzufügen
	body.AddChild(pElement)

	// Aktualisierte XML-Datei speichern
	if err := doc.WriteToFile(docxPath); err != nil {
		fmt.Println("Error when writing the updated XML file:", err)
		return
	}

	fmt.Println("New paragraph was successfully added to <w:body> node.")
}

// Dateisystem

func unzip(zipFilePath, targetFolder string) {
	dst := targetFolder
	archive, err := zip.OpenReader(zipFilePath)
	if err != nil {
		panic(err)
	}
	defer archive.Close()

	for _, f := range archive.File {
		filePath := filepath.Join(dst, f.Name)
		fmt.Println("unzipping file ", filePath)

		if !strings.HasPrefix(filePath, filepath.Clean(dst)+string(os.PathSeparator)) {
			fmt.Println("invalid file path")
			return
		}
		if f.FileInfo().IsDir() {
			fmt.Println("creating directory...")
			os.MkdirAll(filePath, os.ModePerm)
			continue
		}

		if err := os.MkdirAll(filepath.Dir(filePath), os.ModePerm); err != nil {
			panic(err)
		}

		dstFile, err := os.OpenFile(filePath, os.O_WRONLY|os.O_CREATE|os.O_TRUNC, f.Mode())
		if err != nil {
			panic(err)
		}

		fileInArchive, err := f.Open()
		if err != nil {
			panic(err)
		}

		if _, err := io.Copy(dstFile, fileInArchive); err != nil {
			panic(err)
		}

		dstFile.Close()
		fileInArchive.Close()
	}
}

func zipFolder(folderPath, zipFileName string) {
	// Ordnerpfad, den du zippen möchtest
	sourceFolder := folderPath

	// Erstelle eine neue ZIP-Datei
	zipFile, err := os.Create(zipFileName)
	if err != nil {
		panic(err)
	}
	defer zipFile.Close()

	// Erstelle einen neuen ZIP-Writer
	zipWriter := zip.NewWriter(zipFile)
	defer zipWriter.Close()

	// Gehe durch alle Dateien im Quellordner
	err = filepath.Walk(sourceFolder, func(path string, info os.FileInfo, err error) error {
		if err != nil {
			return err
		}

		// Überspringe Ordner
		if info.IsDir() {
			return nil
		}

		// Öffne die Datei
		file, err := os.Open(path)
		if err != nil {
			return err
		}
		defer file.Close()

		// Erstelle einen neuen ZIP-Dateieintrag
		relPath, err := filepath.Rel(sourceFolder, path)
		if err != nil {
			return err
		}
		zipEntry, err := zipWriter.Create(relPath)
		if err != nil {
			return err
		}

		// Kopiere den Inhalt der Datei in den ZIP-Eintrag
		_, err = io.Copy(zipEntry, file)
		if err != nil {
			return err
		}

		return nil
	})

	if err != nil {
		fmt.Println("Error when creating the ZIP file:", err)
	} else {
		fmt.Println("ZIP file successfully created:", zipFileName)
	}
}

func deleteFolder(folderPath string) error {
	// Ordner löschen
	err := os.RemoveAll(folderPath)
	if err != nil {
		return err
	}
	fmt.Printf("Folder '%s' successfully deleted.\n", folderPath)
	return nil
}
