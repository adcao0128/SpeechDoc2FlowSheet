<template>
  <div class="hello">
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
      <label for="WordDocDownload">Choose Speech Document</label> <br />
      <input type="file" id="WordDocDownload" name="WordDocDownload" @change="uploadFile" ref="file"/><br />
      <label for="FlowUpdate">Update Flow</label> <br />
      <input type="file" id="FlowUpdate" name="FlowUpdate"  @change="uploadSheet" ref="flow"/><br />
      <label for="headingTypes">Insert the types of headings you use: For example, "h1, h2, h3"</label> <br />
      <input type="text" id="headingTypes" name="headingTypes" @change="changeHeading($event)" value="h1, h2, h3"><br />
      <button v-on:click="submitFiles">Submit</button> <br />
      <b class="center">
        Remember to use the wrap text function in the Excel Document!
      </b>
  </div>
</template>
<!--a :href="this.newSheet" download>Download</a> -->
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
    sleep(ms) {
      return new Promise(resolve => setTimeout(resolve, ms));
    },
    async makeInvisible() {
      await this.sleep(2000);
      const htmlDiv = document.getElementById("A");
      htmlDiv.style.display = "none";
    },
    async createExcel() {
      let workbook;
      let XLSX = require("xlsx");
      if(this.flow == null) {
        workbook = XLSX.utils.book_new();
      } else {
        const data = await this.flow.arrayBuffer();
        workbook = XLSX.read(data);
      }
      const flowNames = document.querySelectorAll(`${this.heading}`);

      for(let i = 0; i < flowNames.length; i++) {
        let name = flowNames[i];
        let nextName;
        let second2Last = null;
        if(i != flowNames.length - 1) {
          nextName = flowNames[i + 1];
        } else {
          second2Last = flowNames[i - 1];
        }
        let tags = [];
        for (let currentElement of document.querySelectorAll("h4")) {
          if((i != flowNames.length - 1) && currentElement.tagName == "H4" && (currentElement.compareDocumentPosition(name) == 2 && currentElement.compareDocumentPosition(nextName) == 4)) {
            tags.push([currentElement.textContent]);
            tags.push([])
          }
          if((i == flowNames.length - 1) && currentElement.tagName == "H4" && currentElement.compareDocumentPosition(second2Last) == 2) {
            tags.push([currentElement.textContent]);
            tags.push([]);
          }
        }
        if(tags.length != 0) {
          if(!workbook.SheetNames.some(thisName=> name.textContent == thisName)){
            XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(tags, { origin: `${this.speech}2` }), name.textContent);
          } else {
            XLSX.utils.sheet_add_aoa(workbook.Sheets[name.textContent], tags, { origin: `${this.speech}2` })
          }
        }

      }
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
      

      var data = XLSX.writeFile(workbook, "flow.xlsx");
      this.newSheet = data;   
      await this.sleep(2000);
    },
    async submitFiles() {
      require("docx2html")(this.word)
            .then(html=>{
                html.toString()
            });
      this.makeInvisible();
      await this.sleep(2000);
      this.createExcel();
      console.log(this.speech);
    }
  }

  
}
</script>

<!-- Add "scoped" attribute to limit CSS to this component only -->
<style scoped>
h3 {
  margin: 40px 0 0;
}
ul {
  list-style-type: none;
  padding: 0;
}
li {
  display: inline-block;
  margin: 0 10px;
}
a {
  color: #42b983;
}
</style>
