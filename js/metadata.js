/**
 * 获取文档元信息
 */
// 
// 标记各值是否取得
// 0b0001 文号
// 0b0010 标题
// 0b0100 主送
// 0b1000 发文日期

const HAS_CODE = 0b0001;
const HAS_TITLE = 0b0010;
const HAS_SEND_TO = 0b0100;
const HAS_SEND_DATE = 0b1000;


function copyToClipboard(val) {
    const copyText = document.createElement('input')
    copyText.value = val
    document.body.appendChild(copyText)
    copyText.focus()
    copyText.select()
    try {
        const successful = document.execCommand('copy')
        const msg = successful ? 'successful' : 'unsuccessful'
        console.log('Copying text command was ' + msg)
    } catch (err) {
        console.log('Oops, unable to copy')
    }
}


function searchRegex(val, re) {
    const mat = val.replace(' ', '').replace('　', '').match(re)
    return mat ? mat[0] : null
}

function searchCode(val) {
    const re = /^\S+〔\d{4}〕\d+号$/
    return searchRegex(val, re)
}

function searchSendTo(val) {
    const re = /^\S+[：:]$/
    return searchRegex(val, re)
}

function searchSendDate(val) {
    const re = /^\d{4}年\d{1,2}月\d{1,2}日$/
    return searchRegex(val, re)
}


function getLines() {
    const doc = wps.WpsApplication().ActiveDocument
    const paragraphs = doc.Paragraphs
    const lines = []
    for (let index = 1; index <= paragraphs.Count; index++) {
        const p = paragraphs.Item(index).Range.Text.trim();
        if (p !== '') {
            lines.push(p)
        }
    }
    return lines
}


function prefixInteger(num, length) {
    return ('0000000000000000' + num).substr(-length);
}

function parseDoc() {
    const meta = {}
    const lines = getLines()
    let flag = 0b0000;
    let codeLineNum = 0

    for (const line of lines) {
        // 如果没有文号，没有标题，搜索文号
        if ((flag & HAS_CODE) === 0 && (flag & HAS_TITLE) === 0) {
            const code = searchCode(line)
            if (code) {
                meta.code = code
                codeLineNum = lines.indexOf(line)
                flag |= HAS_CODE
                continue
            }
        }

        // 如果没有主送，搜索主送，并获得标题
        if ((flag & HAS_SEND_TO) === 0) {
            const sendTo = searchSendTo(line)
            if (sendTo) {
                meta.sendTo = sendTo
                flag |= HAS_SEND_TO
                const sendToLineNum = lines.indexOf(line)
                const titleArray = lines.slice(codeLineNum + 1, sendToLineNum)
                meta.title = titleArray.join('')
                flag |= HAS_TITLE
                continue
            }
        }

        // 如果没有日期，搜索日期，并获得发文单位
        if ((flag & HAS_SEND_DATE) === 0) {
            const sendDate = searchSendDate(line)

            if (sendDate) {
                const sendDateLineNum = lines.indexOf(line)
                console.log('da', sendDate)
                meta.sendBy = lines[sendDateLineNum - 1]
                meta.sendDate = sendDate
                flag |= HAS_SEND_DATE
                continue
            }
        }

        if (flag === (HAS_CODE | HAS_TITLE | HAS_SEND_TO | HAS_SEND_DATE)) {
            break
        }
    }

    console.log('flag', `0b${prefixInteger(flag.toString(2), 4)}`)
    console.log('meta', meta)
    return meta
}

function lineStringify(obj) {
    return `${obj.title}\t${obj.code}`
}

function getMetaData(type = 'object') {
    const meta = parseDoc()
    const result = type === 'object' ? JSON.stringify(meta) : lineStringify(meta)
    copyToClipboard(result)
    alert('成功复制到剪贴板')
}
