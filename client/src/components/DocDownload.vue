<template>
  <div>
    <b class="center">
      Place a Speech Document and your current flow to create a new flow
      document
    </b>
    <br />

      <label for="1AC">1AC:</label>
      <input type="radio" id="1AC" name="speech" @change="speechChange($event)" value="A" checked/><br />
      <label for="1NC">1NC:</label>
      <input type="radio" id="1NC" name="speech" @change="speechChange($event)" value="B"/><br />
      <label for="2AC">2AC:</label>
      <input type="radio" id="2AC" name="speech" @change="speechChange($event)" value="C"/><br />
      <label for="2NC">2NC:</label>
      <input type="radio" id="2NC" name="speech" @change="speechChange($event)" value="D"/><br />
      <label for="1NR">1NR:</label>
      <input type="radio" id="1NR" name="speech" @change="speechChange($event)" value="E"/><br />
      <label for="1AR">1AR:</label>
      <input type="radio" id="1AR" name="speech" @change="speechChange($event)" value="F"/><br />
      <label for="2NR">2NR:</label>
      <input type="radio" id="2NR" name="speech" @change="speechChange($event)" value="G"/><br />
      <label for="2AR">2AR:</label>
      <input type="radio" id="2AR" name="speech" @change="speechChange($event)" value="H"/><br />
      <label for="WordDocDownload">Choose Speech Document</label>
      <input type="file" id="WordDocDownload" name="WordDocDownload" @change="uploadFile" ref="file"/><br />
      <label for="FlowUpdate">Update Flow</label>
      <input type="file" id="FlowUpdate" name="FlowUpdate"  @change="uploadSheet" ref="flow"/><br />
      <label for="headingTypes">Insert the types of headings you use: For example, "h1, h2, h3"</label> <br />
      <input type="text" id="headingTypes" name="headingTypes" @change="changeHeading($event)" value="h1, h2, h3"><br />
      <button v-on:click="submitFiles">Submit</button> <br />
      <b class="center">
        Remember to use the wrap text function in the Excel Document!
      </b>
      <br />
      <br />
      <br />
      <b id="SubmittedSpeechDoc" class="center" style="visibility: hidden">
        Below is the Speech Document you uploaded.
      </b>
  </div>
</template>
<script>

export default {
  name: 'DocDownload',
  data() {
    return {
      word: null,
      flow: null,
      newSheet: null,
      speech: "A",
      heading: "h1, h2, h3",
    }
  },
  methods: {
    uploadSheet() {
      this.flow = this.$refs.flow.files[0];
    },
    changeHeading(event) {
      this.heading = event.target.value;
    },
    speechChange(event) {
      this.speech = event.target.value;
    },
    uploadFile() {
      this.word = this.$refs.file.files[0];
    },
    //creates the excel sheet
    async createExcel() {
      let workbook;
      let XLSX = require("xlsx");
      //Either creates a new workbook or gets a workbook from the given sheet
      if(this.flow == null) {
        workbook = XLSX.utils.book_new();
      } else {
        const data = await this.flow.arrayBuffer();
        workbook = XLSX.read(data);
      }
      //Grabs the headings for the flow sheets
      const flowNames = document.querySelectorAll(`${this.heading}`);
      //Get the tags for each flow sheet
      for(let i = 0; i < flowNames.length; i++) {
        //Sets up which tags to take from heading to heading
        let name = flowNames[i];
        let nextName;
        let second2Last = null;
        if(i != flowNames.length - 1) {
          nextName = flowNames[i + 1];
        } else {
          second2Last = flowNames[i - 1];
        }
        let tags = [];
        //Goes through the entire HTML dom
        for (let currentElement of document.querySelectorAll("h4")) {
          //Either we are at the start/middle of the document
          if((i != flowNames.length - 1) && currentElement.tagName == "H4" && (currentElement.compareDocumentPosition(name) == 2 && currentElement.compareDocumentPosition(nextName) == 4)) {
            tags.push([currentElement.textContent]);
            tags.push([])
          }
          //Or the last heading for the document
          if((i == flowNames.length - 1) && currentElement.tagName == "H4" && currentElement.compareDocumentPosition(second2Last) == 2) {
            tags.push([currentElement.textContent]);
            tags.push([]);
          }
        }
        //Limit sheetnames to 31 characters and remove bad input
        if(name.textContent.length > 31) {
          name.textContent = name.textContent.slice(0, 31).replaceAll(":", "|").replaceAll("/", "|").replaceAll("\\", "|");
        } else {
          name.textContent = name.textContent.replaceAll(":", "|").replaceAll("/", "|").replaceAll("\\", "|");
        }
        //Tags are added to a new sheet, or appended to a current one.
        if(tags.length != 0) {
          if(!workbook.SheetNames.some(thisName=> name.textContent == thisName)){
            XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(tags, { origin: `${this.speech}2` }), name.textContent);
          } else {
            XLSX.utils.sheet_add_aoa(workbook.Sheets[name.textContent], tags, { origin: `${this.speech}2` })
          }
        }

      }
      //Sets each sheet to be organized by speech, and create a nice columnn width
      workbook.SheetNames.forEach(name => {
        let worksheet = workbook.Sheets[name];
        let wscols = [
          { width: 25 },
          { width: 25 },
          { width: 25 }, 
          { width: 25 }, 
          { width: 25 },
          { width: 25 }, 
          { width: 25 }, 
          { width: 25 }, 
          { width: 25 },
          { width: 25 }
        ];
        worksheet["!cols"] = wscols;
        XLSX.utils.sheet_add_aoa(worksheet, [["1AC", "1NC", "2AC", "2NC", "1NR", "1AR", "2NR", "2AR"]], { origin: "A1" });
      });
      
      //Download the excel sheet
      var data = XLSX.writeFile(workbook, "flow.xlsx");
      this.newSheet = data;   
    },
    async submitFiles() {
      await require("docx2html")(this.word)
            .then(html=>{
                html.toString()
            });
      document.querySelector("#SubmittedSpeechDoc").style.visibility="visible";
      this.createExcel();
    }
  }

  
}
</script>
