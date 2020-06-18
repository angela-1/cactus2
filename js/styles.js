/**
 * 套用公文格式
 */

const WdOutlineLevel = {
    WdOutlineLevel1: 1,
    WdOutlineLevel2: 2,
    WdOutlineLevel10: 10
}

function __getSelectionStart(selection, pars) {
    let startPar = 1;
    for (let i = 1; i <= pars.Count; i++) {
        const p = pars.Item(i)
        if (p.Range.End >= selection.Start) {
            startPar = i
            break
        }
    }
    return startPar
}

function __getSelectionEnd(selection, pars) {
    let endPar = 1;
    for (let i = 1; i <= pars.Count; i++) {
        const p = pars.Item(i);
        if (p.Range.Start >= selection.End) {
            endPar = i - 1
            break
        }
    }
    return endPar
}

function __regSearch(re, row) {
    const mat = row.match(re)
    return !!mat
}


function __setOfficeStyles() {
    const app = wps.WpsApplication()
    const doc = app.ActiveDocument
    doc.Content.ListFormat.ConvertNumbersToText()
    app.Selection.ClearFormatting()

    const pars = doc.Paragraphs
    const selection = app.Selection

    if (selection.Paragraphs.Count < 2) {
        alert('请选中要套用格式的段落')
        return
    }

    const selectionPars = selection.Paragraphs
    const selectionCount = selectionPars.Count

    const startIndex = __getSelectionStart(selection, pars)
    const endIndex = __getSelectionEnd(selection, pars)

    console.log('star se', startIndex, endIndex)

    let lockLevel1Font = false
    let lockLevel2Font = false

    const re1 = /^[一二三四五六七八九十]+、/g
    const re2 = /^（[一二三四五六七八九十]+）/g

    for (let i = startIndex; i <= endIndex; i++) {
        console.log('see', i, selectionCount, pars.Item(i).Range.Text)
        pars.Item(i).Range.Text = pars.Item(i).Range.Text
            .replace(new RegExp(' ', 'g'), '').replace(new RegExp('　', 'g'), '')
        const text = pars.Item(i).Range.Text

        if (__regSearch(re1, text)) {
            pars.Item(i).Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.WdOutlineLevel1
            if (text.Length > 24) {
                lockLevel1Font = true
            }
            console.log('l 1', pars.Item(i).Range.ParagraphFormat.OutlineLevel, text);
        } else if (__regSearch(re2, text)) {
            pars.Item(i).Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.WdOutlineLevel2
            if (text.Length > 24) {
                lockLevel2Font = true
            }
            console.log('l 2', pars.Item(i).Range.ParagraphFormat.OutlineLevel, text);

        } else {
            pars.Item(i).Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.WdOutlineLevel10
            console.log('p', i, pars.Item(i).Range.ParagraphFormat.OutlineLevel, text);
        }
    }

    for (let i = startIndex; i <= endIndex; i++) {
        pars.Item(i).LeftIndent = app.CentimetersToPoints(0);
        pars.Item(i).RightIndent = app.CentimetersToPoints(0);
        pars.Item(i).SpaceBefore = 0;
        pars.Item(i).SpaceBeforeAuto = 0;
        pars.Item(i).SpaceAfter = 0;
        pars.Item(i).SpaceAfterAuto = 0;
        pars.Item(i).CharacterUnitLeftIndent = 0;
        pars.Item(i).CharacterUnitRightIndent = 0;
        pars.Item(i).CharacterUnitFirstLineIndent = 2;
        pars.Item(i).Range.Font.Size = 16;
        pars.Item(i).Range.Font.Bold = 0;
        pars.Item(i).LineUnitBefore = 0;
        pars.Item(i).LineUnitAfter = 0;

        // p.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;
        // p.LineSpacing = 28;

        const level = pars.Item(i).Range.ParagraphFormat.OutlineLevel
        console.log('sha', i, level, pars.Item(i).Range.Text);

        switch (level) {
            case WdOutlineLevel.WdOutlineLevel1:
                if (lockLevel1Font) {
                    pars.Item(i).Range.Font.Name = '方正仿宋_GBK'
                } else {
                    pars.Item(i).Range.Font.Name = '方正黑体_GBK'
                }
                pars.Item(i).Range.Font.NameAscii = 'Times New Roman'
                break
            case WdOutlineLevel.WdOutlineLevel2:
                if (lockLevel2Font) {
                    pars.Item(i).Range.Font.Name = '方正仿宋_GBK'
                } else {
                    pars.Item(i).Range.Font.Name = '方正楷体_GBK'
                }
                pars.Item(i).Range.Font.NameAscii = 'Times New Roman'
                console.log('level 2', pars.Item(i).Range.Font.NameAscii, pars.Item(i).Range.Text);

                break

            default:
                pars.Item(i).Range.Font.Name = '方正仿宋_GBK'
                pars.Item(i).Range.Font.NameAscii = 'Times New Roman'
                pars.Item(i).Range.ParagraphFormat.WidowControl = 0; // 不勾选 孤行控制 
                console.log('zheng wen', pars.Item(i).Range.Font.NameAscii, pars.Item(i).Range.Text);

                break
        }
    }

    selection.SetRange(0, 0)
    alert('套用格式完成')
}


function setStyles() {
    __setOfficeStyles()
}