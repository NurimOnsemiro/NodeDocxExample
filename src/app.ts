//https://www.npmjs.com/package/docx

import {
    Alignment,
    AlignmentType,
    Document,
    HeadingLevel,
    HeightRule,
    Media,
    Packer,
    Paragraph,
    Table,
    TableCell,
    TableRow,
    VerticalAlign,
    WidthType,
} from 'docx';
import fs from 'fs';

const doc = new Document();

const image1 = Media.addImage(doc, fs.readFileSync('./resources/frame.png'));

function makeRowTemplate1(title: string, subTitle: string = ''): TableRow {
    let row = new TableRow({
        height: {
            height: 480,
            rule: HeightRule.EXACT,
        },
        children: [
            new TableCell({
                verticalAlign: VerticalAlign.CENTER,
                width: {
                    size: 2115,
                    type: WidthType.DXA,
                },
                children: [
                    new Paragraph({
                        text: title,
                        alignment: AlignmentType.CENTER,
                    }),
                ],
            }),
            new TableCell({
                verticalAlign: VerticalAlign.CENTER,
                width: {
                    size: 7350,
                    type: WidthType.DXA,
                },
                children: [
                    new Paragraph({
                        text: subTitle,
                        alignment: AlignmentType.LEFT,
                    }),
                ],
            }),
        ],
    });

    return row;
}

const table = new Table({
    rows: [
        makeRowTemplate1('납 품 처'),
        makeRowTemplate1('계 약 명'),
        makeRowTemplate1('MAC ADD.', '(기존 사이트 생략가능)'),
        new TableRow({
            children: [
                new TableCell({
                    children: [
                        new Paragraph({
                            text: 'World',
                            heading: HeadingLevel.HEADING_1,
                        }),
                    ],
                }),
                new TableCell({
                    children: [new Paragraph(image1)],
                }),
            ],
        }),
    ],
});

doc.addSection({
    children: [
        new Paragraph({
            text: '라이선스 신청서',
            heading: HeadingLevel.HEADING_1,
            alignment: AlignmentType.CENTER,
            border: {
                bottom: {
                    color: 'black',
                    value: 'single',
                    size: 6,
                    space: 1,
                },
            },
        }),
        table,
        new Paragraph(image1),
    ],
});

Packer.toBuffer(doc).then(buffer => {
    fs.writeFileSync('MyDocument.docx', buffer);
});
