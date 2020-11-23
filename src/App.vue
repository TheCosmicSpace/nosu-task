<template>
  <div id="app">
    <label for="custom-file-upload" class="filupp">
      <span class="filupp-file-name js-value">Выберите файл</span>
      <input @change="onSelectFile" type="file" name="attachment-file" id="custom-file-upload"/>
    </label>
  </div>
</template>

<script>
import xlsxParser from 'xls-parser'
export default {
  methods: {
    onSelectFile (event) {
      this.file = event.target.files ? event.target.files[0] : null
      console.log(this.file)
      if (!this.file) return alert('Выберите файл')
      this.enterSearchLesson()
      this.parseSelectedFile()
    },
    enterSearchLesson () {
      this.searchLesson = prompt('Введите название предмета')
      if (!this.searchLesson) this.enterSearchLesson()
    },
    parseSelectedFile () {
      xlsxParser
        .onFileSelection(this.file)
        .then(data => {
          const parsedData = data
          this.chechSheet(parsedData)
        })
    },
    chechSheet (parsedData) {
      let competencies = []
      for (const [key, value] of Object.entries(parsedData)) {
        if (key.match(/компетенции/gi)) {
          competencies = [...competencies, ...new Set(this.constructCompetencies(value))]
        }
      }
      this.generateWordDocument(competencies)
    },
    constructCompetencies (collection) {
      return collection.reduce((acc, lesson) => {
        const lessonName = lesson['Содержание'] || lesson['Наименование'] || ''
        if (lessonName.toLowerCase() === this.searchLesson.toLowerCase() &&
            lesson['Формируемые компетенции']) acc.push(lesson['Формируемые компетенции'])
        return acc
      }, [])
    },
    generateWordDocument (competencies) {
      if (!competencies.length) return alert('Ничего не найдено')
      const doc = new window.docx.Document()
      const children = competencies.map(item => {
        return new window.docx.TextRun(item)
      })
      doc.addSection({
        properties: {},
        children: [
          new window.docx.Paragraph({
            children
          })
        ]
      })
      this.saveDocumentToFile(doc, this.searchLesson)
    },
    saveDocumentToFile (doc, fileName) {
      window.docx.Packer.toBlob(doc).then(blob => {
        window.saveAs(blob, fileName)
      })
      this.resetData()
    },
    resetData () {
      this.file = null
      this.searchLesson = ''
      this.competencies = []
    }
  },
  data: () => ({
    file: null,
    searchLesson: '',
    competencies: []
  }),
  name: 'App'
}
</script>

<style>
#app {
  font-family: Avenir, Helvetica, Arial, sans-serif;
  -webkit-font-smoothing: antialiased;
  -moz-osx-font-smoothing: grayscale;
  text-align: center;
  color: #2c3e50;
  margin-top: 60px;
}
html {
  box-sizing: border-box;
  background: #181818;
  overflow-y: scroll;
}
body {
  position: relative;
  font: 1em/1.6 "Varela Round", Arial, sans-serif;
  color: #999;
  font-weight: 400;
  max-width: 25em;
  padding: 1em;
  margin: 10% auto;
}

*, *:before, *:after {
  box-sizing: inherit;
}

h2 {
  font-weight: 400;
}

.filupp > input[type="file"] {
  position: absolute;
  width: 1px;
  height: 1px;
  padding: 0;
  margin: -1px;
  overflow: hidden;
  clip: rect(0,0,0,0);
  border: 0;
}

.filupp {
  position: relative;
  background: #242424;
  display: block;
  padding: 1em;
  font-size: 1em;
  width: 100%;
  height: 3.5em;
  color: #fff;
  cursor: pointer;
  box-shadow: 0 1px 3px darken(#242424,10);
}

.filupp:before {
  content: "";
  position: absolute;
  top: 1.5em;
  right: .75em;
  width: 2em;
  height: 1.25em;
  border: 3px solid #dd4040;
  border-top: 0;
  text-align: center;
}

.filupp:after {
  content: "\279c";
  transform: rotate(-90deg);
  position: absolute;
  top: .65em;
  right: .45em;
  font-size: 2em;
  color: #dd4040;
  line-height: 0;
}

.filupp-file-name {
  width: 75%;
}
</style>
