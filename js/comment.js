/**
 * 导出批注
 */

const WdInformation = {
    wdActiveEndPageNumber: 3,
    wdFirstCharacterLineNumber: 10
}


function parseComments() {
    const doc = wps.WpsApplication().ActiveDocument
    const fileName = doc.Name;
    const comments = doc.Comments
    if (comments.Count === 0) {
        alert('此文档中没有批注')
        return
    }
    const res = [fileName]
    for (let i = 1; i < comments.Count + 1; i++) {
        const c = comments.Item(i)
        console.log('cc', c);

        const commentObject = {
            pagenumber: c.Scope.Information(WdInformation.wdActiveEndPageNumber),
            linenumber: c.Scope.Information(WdInformation.wdFirstCharacterLineNumber),
            src: c.Scope.Text,
            comment: c.Range.Text,
            author: c.Author
        }
        res.push(commentObject)
    }
    return res
}

function writeToDoc(comments) {
    const newDoc = wps.WpsApplication().Documents.Add()
    newDoc.Content.ParagraphFormat.CharacterUnitFirstLineIndent = 2;
    newDoc.Content.Paragraphs.Item(1).Range.Font.Size = 16;
    newDoc.Content.Paragraphs.Item(1).Range.Font.Name = '方正仿宋_GBK';
    newDoc.Content.Paragraphs.Item(1).Range.Font.NameAscii = 'Times New Roman';

    const par = newDoc.Content.Paragraphs.Add();
    par.Range.Text = '导出批注工具';
    par.Range.InsertParagraphAfter();

    const fileName = comments.shift()
    par.Range.InsertAfter('以下是来自文档“' + fileName + '”的批注。');
    par.Range.InsertParagraphAfter();

    par.Range.InsertAfter('------------------------分割线------------------------');
    par.Range.InsertParagraphAfter();

    for (let i = 0; i < comments.length; i++) {
        const element = comments[i];
        par.Range.InsertAfter(`${i + 1}、第${element.pagenumber}页，第${element.linenumber}行 || ` +
            `原文：${element.src} || 意见：${element.comment}`);
    }
}


function getComments() {
    const comments = parseComments()
    writeToDoc(comments)
}