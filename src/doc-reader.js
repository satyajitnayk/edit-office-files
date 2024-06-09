import fs from 'fs.js';
import JSZip from 'jszip';
import xml2js from 'xml2js';

export class DocxReader {
  constructor(filename) {
    this.filename = filename;
  }

  async loadZip() {
    try {
      // Read the .docx file as binary
      const data = fs.readFileSync(this.filename);
      // Initialize JSZip and load the data
      this.zip = await JSZip.loadAsync(data);
    } catch (error) {
      console.error("Error reading .docx file:", error);
      throw error;
    }
  }

  async readString(filePath) {
    try {
      if (!this.zip) {
        await this.loadZip();
      }
      // Extract the specified file from the ZIP
      return await this.zip.file(filePath).async("string");
    } catch (error) {
      console.error(`Error extracting file "${filePath}":`, error);
      throw error;
    }
  }

  async readAsObject(filePath) {
    try {
      const fileContent = await this.readString(filePath);
      // Parse the XML content
      const parser = new xml2js.Parser();
      return await parser.parseStringPromise(fileContent);
    } catch (error) {
      console.error("Error parsing XML content:", error);
      throw error;
    }
  }

  async writeString(filePath, content) {
    try {
      if (!this.zip) {
        await this.loadZip();
      }
      // Write the string content to the specified file in the ZIP
      this.zip.file(filePath, content);
    } catch (error) {
      console.error(`Error writing string to file "${filePath}":`, error);
      throw error;
    }
  }

  async writeObject(filePath, content) {
    try {
      // Convert the object to an XML string
      const builder = new xml2js.Builder();
      const xmlContent = builder.buildObject(content);
      // Write the XML string to the specified file in the ZIP
      await this.writeString(filePath, xmlContent);
    } catch (error) {
      console.error("Error writing object to XML file:", error);
      throw error;
    }
  }

  async save(filename = '') {
    try {
      if (!this.zip) {
        await this.loadZip();
      }
      if (!filename) {
        filename = this.filename;
      }
      // Generate the new .docx file as a binary
      const newDocx = await this.zip.generateAsync({type: 'nodebuffer'});
      // Write the binary to a new file
      fs.writeFileSync(filename, newDocx);
    } catch (error) {
      console.error("Error saving .docx file:", error);
      throw error;
    }
  }
}
