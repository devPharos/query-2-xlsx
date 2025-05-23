import Sequelize from 'sequelize'
import { resolve } from 'path'
import { Connection, Request } from 'tedious'
import { config } from '../../config/database'

const xl = require('excel4node')
const fs = require('fs')

function unicodeToChar(text) {
    return text.replace(/\\u[\dA-F]{4}/gi, function (match) {
        return String.fromCharCode(parseInt(match.replace(/\\u/g, ''), 16))
    })
}

async function geraExcel({
    res = null,
    Resultados = [],
    Nome = 'default_name',
    Titulos = [],
    Formatos = [],
    Tamanhos = [],
}) {
    const name = `${Nome}_${Date.now()}`
    const path = `${resolve(
        __dirname,
        '..',
        '..',
        '..',
        'public',
        'uploads'
    )}/${name}.xlsx`

    const wb = new xl.Workbook()
    // Add Worksheets to the workbook
    var ws = wb.addWorksheet('Dados')

    // Create a reusable style
    var styleBold = wb.createStyle({
        font: {
            color: '#222222',
            size: 12,
            bold: true,
        },
    })

    var styleTotal = wb.createStyle({
        font: {
            color: '#00aa00',
            size: 12,
            bold: true,
        },
        fill: {
            type: 'pattern',
            fgColor: '#ff0000',
            bgColor: '#ffffff',
        },
        numberFormat: '$ #,##0.00; ($#,##0.00); -',
    })

    var styleTotalNegative = wb.createStyle({
        font: {
            color: '#aa0000',
            size: 12,
            bold: true,
        },
        fill: {
            type: 'pattern',
            fgColor: '#ff0000',
            bgColor: '#ffffff',
        },
        numberFormat: '$ #,##0.00; ($#,##0.00); -',
    })

    var styleHeading = wb.createStyle({
        font: {
            color: '#222222',
            size: 14,
            bold: true,
        },
        alignment: {
            horizontal: 'center',
            vertical: 'center',
        },
    })

    var styleBody = wb.createStyle({
        font: {
            size: 12,
        },
    })

    var styleBodyCenter = wb.createStyle({
        font: {
            size: 12,
        },
        alignment: {
            horizontal: 'center',
            vertical: 'center',
        },
        numberFormat: 'dd/mm/yyyy',
    })

    var styleBodyRight = wb.createStyle({
        font: {
            size: 12,
        },
        alignment: {
            horizontal: 'right',
            vertical: 'center',
        },
    })

    var styleNumeric = wb.createStyle({
        font: {
            size: 12,
        },
        alignment: {
            horizontal: 'center',
            vertical: 'center',
        },
        numberFormat: '#,##0.00;[Red]-#,##0.00; -',
    })

    var styleCurrency = wb.createStyle({
        font: {
            size: 12,
        },
        alignment: {
            horizontal: 'center',
            vertical: 'center',
        },
        numberFormat: 'R$ #,##0.00;[Red]-R$ #,##0.00; -',
    })

    if (Resultados.length === 0) {
        const err = 'Registros não encontrados'
        const stats = 'Registros não encontrados'
        const ret = res.status(400).json({ err, stats })
        return ret
    }

    let row = 1

    Titulos.forEach((titulo, index) => {
        ws.cell(row, index + 1)
            .string(unicodeToChar(titulo))
            .style(styleHeading)
    })

    Tamanhos.forEach((tamanho, index) => {
        ws.column(index + 1).setWidth(tamanho)
    })

    Resultados.forEach((rows) => {
        ++row
        rows.forEach((valor, index) => {
            if (Formatos[index] === 1) {
                ws.cell(row, index + 1)
                    .string(valor || '')
                    .style(styleBody)
            } else if (Formatos[index] === 4) {
                if (!valor) {
                    ws.cell(row, index + 1).string('')
                } else if (typeof valor === 'string' && valor.includes('/')) {
                    const dataArray = valor.split('/')
                    const dataValor =
                        dataArray[2] + '-' + dataArray[1] + '-' + dataArray[0]
                    ws.cell(row, index + 1)
                        .date(dataValor)
                        .style(styleBodyCenter)
                } else if (typeof valor === 'string' && !valor.includes('-')) {
                    const dataValor =
                        valor.substring(0, 4) +
                        '-' +
                        valor.substring(4, 6) +
                        '-' +
                        valor.substring(6, 8)
                    ws.cell(row, index + 1)
                        .date(dataValor)
                        .style(styleBodyCenter)
                } else {
                    ws.cell(row, index + 1)
                        .date(valor)
                        .style(styleBodyRight)
                }
            } else if (Formatos[index] === 3) {
                ws.cell(row, index + 1)
                    .number(valor || 0)
                    .style(styleCurrency)
            } else {
                ws.cell(row, index + 1)
                    .number(valor || 0)
                    .style(styleNumeric)
            }
        })
    })

    ws.row(1).filter()
    ws.row(1).freeze()

    let ret = null
    wb.write(path, (err, stats) => {
        if (err) {
            ret = res.status(400).json({ err, stats })
        } else {
            setTimeout(() => {
                fs.unlink(path, (err) => {
                    if (err) {
                        console.log(err)
                    }
                })
            }, process.env.TEMPO_PARA_EXCLUSAO || 360000)
            return res.json({ path, name })
        }
    })
    return ret
}

const { Op } = Sequelize
class GenerateExcelController {
    async store(req, res) {
        const { Query, Nome, Titulos, Formatos, Tamanhos } = req.body

        var conn = new Connection(config)
        conn.on('connect', function (err) {
            const request = new Request(Query, function (err) {
                if (err) {
                    console.log(err)
                }
            })
            const Resultados = []
            request.on('row', function (columns) {
                const row = []
                columns.forEach(function (column, index) {
                    row.push(column.value)
                })
                Resultados.push(row)
            })

            // Close the connection after the final event emitted by the request, after the callback passes
            request.on('requestCompleted', async function (rowCount, more) {
                conn.close()
                await geraExcel({
                    res,
                    Resultados,
                    Nome,
                    Titulos,
                    Formatos,
                    Tamanhos,
                })
            })

            conn.execSql(request)
        })

        conn.connect()
    }
}

export default new GenerateExcelController()
