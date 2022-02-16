import fs from "fs/promises"
import neatCsv from "neat-csv"
import yargs from "yargs"
import pptxgen from "pptxgenjs"

const main = async () => {
  const parser = yargs(process.argv.slice(2)).options({
    filename: { type: "string", demandOption: true },
  })
  // .help()
  // .parse()

  const argv = await parser.argv
  console.log(argv.filename)

  const file = await fs.readFile(argv.filename)
  const rows = await neatCsv(file)
  const discipline = "Bouldering"
  const climbers = rows.map((row, i) => {
    const num = i + 1
    const {
      Name: name,
      Category: category,
      RO: order,
      Bib: bib,
      Time: time,
      Team: team,
    } = row

    return { name, category, discipline, order, bib, time, team, num }
  })
  const numClimbers = climbers.length
  console.log(`${climbers.length} climbers`)
  const includeNum = false
  const pres = new pptxgen()
  climbers.forEach((climber) => {
    const slide = pres.addSlide()

    slide.addText(climber.category, {
      x: "10%",
      y: "20%",
      w: "80%",
      align: pres.AlignH.right,
      valign: pres.AlignV.middle,
      color: "#A9A9A9",
      fontSize: 36,
    })

    slide.addText(climber.discipline, {
      x: "10%",
      y: "20%",
      w: "80%",
      align: pres.AlignH.left,
      valign: pres.AlignV.middle,
      color: "#A9A9A9",
      fontSize: 36,
    })

    slide.addText(climber.name, {
      x: "10%",
      y: "45%",
      w: "80%",
      align: pres.AlignH.center,
      valign: pres.AlignV.middle,
      color: pres.SchemeColor.text1,
      fontSize: 56,
      bold: true,
    })

    if (climber.team) {
      slide.addText(climber.team, {
        x: "10%",
        y: "70%",
        w: "80%",
        align: pres.AlignH.center,
        valign: pres.AlignV.middle,
        fontSize: 44,
        color: "#A9A9A9",
      })
    }

    if (includeNum) {
      slide.addText(`${climber.num} / ${numClimbers}`, {
        x: "10%",
        y: "85%",
        w: "80%",
        align: pres.AlignH.right,
        valign: pres.AlignV.middle,
        fontSize: 32,
        color: "#A9A9A9",
      })
    }
  })
  const fileName = "output.pptx"
  pres.writeFile({ fileName })
}

main()

// yargs
//   .scriptName("csv2ppt")
//   .usage("$0 <cmd> [args]")
//   .command(
//     "hello [name]",
//     "welcome ter yargs!",
//     (yargs) => {
//       yargs.positional("filename", {
//         type: "string",
//         // default: "Cambi",
//         describe: "Name of CSV file",
//       })
//     },
//     (argv) => {
//       console.log("hello", argv.name, "welcome to yargs!")
//     }
//   )
//   .help().argv
