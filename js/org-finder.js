/**
 * 查找文档里的单位名称
 */


function regSearch(line) {
    const re = /（(牵头|责任|配合)(单位|部门)：(.+?)）/
    const mat = line.match(re)
    return mat ? mat[0] : null
}

function stripBracket(line) {
    const re0 = /[（）\(\) ]/g
    const re1 = /(牵头|责任|配合)(单位|部门)：/g
    const re2 = /[：，。]/g
    line = line.replace(re0, '').replace(re1, '').replace(re2, '、')
    const orgList = line.split('、')
    return orgList
}

function parseOrgs() {
    const doc = wps.WpsApplication().ActiveDocument
    const fileName = doc.Name;
    const paragraphs = doc.Paragraphs
    if (paragraphs.Count === 0) {
        alert('文档是空的')
        return null
    }
    let res = []
    for (let i = 1; i < paragraphs.Count + 1; i++) {
        const item = paragraphs.Item(i)
        const line = item.Range.Text.trim()
        const orgString = regSearch(line)
        if (orgString) {
            const orgArray = stripBracket(orgString)
            res = res.concat(orgArray)
        }
    }
    const setRes = new Set(res)
    const orgs = Array.from(setRes)
    orgs.unshift(fileName)
    return orgs
}

function writeToDoc(orgs) {
    const newDoc = wps.WpsApplication().Documents.Add()
    newDoc.Content.ParagraphFormat.CharacterUnitFirstLineIndent = 2;
    newDoc.Content.Paragraphs.Item(1).Range.Font.Size = 16;
    newDoc.Content.Paragraphs.Item(1).Range.Font.Name = '方正仿宋_GBK';
    newDoc.Content.Paragraphs.Item(1).Range.Font.NameAscii = 'Times New Roman';

    const par = newDoc.Content.Paragraphs.Add();
    par.Range.Text = '查找单位工具';
    par.Range.InsertParagraphAfter();

    const fileName = orgs.shift()
    par.Range.InsertAfter('以下是来自文档“' + fileName + '”的单位：');
    par.Range.InsertParagraphAfter();

    par.Range.InsertAfter('------------------------分割线------------------------');
    par.Range.InsertParagraphAfter();

    const str = orgs.join('、')
    console.log('hb', str);
    par.Range.InsertAfter(str);
}


function getOrgs() {
    const orgs = parseOrgs()
    if (orgs) {
        writeToDoc(orgs)
    }
}