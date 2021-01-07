//https://www.npmjs.com/package/docx

import {
    Alignment,
    AlignmentType,
    Document,
    HeadingLevel,
    HeightRule,
    HorizontalPositionAlign,
    Media,
    Packer,
    Paragraph,
    Table,
    TableCell,
    TableRow,
    TextRun,
    TextWrappingSide,
    TextWrappingType,
    UnderlineType,
    VerticalAlign,
    VerticalPositionAlign,
    VerticalPositionRelativeFrom,
    WidthType,
} from 'docx';
import fs from 'fs';

const doc = new Document();

const logo = Media.addImage(doc, fs.readFileSync('./resources/innodep.jpg'), 197, 100, {
    floating: {
        horizontalPosition: {
            offset: 0,
            align: HorizontalPositionAlign.CENTER,
        },
        verticalPosition: {
            offset: 9105520,
        },
        wrap: {
            type: TextWrappingType.NONE,
            side: TextWrappingSide.BOTH_SIDES,
        },
        margins: {
            bottom: 201440,
        },
    },
});

const background = Media.addImage(doc, fs.readFileSync('./resources/background.png'), 750, 1050, {
    floating: {
        horizontalPosition: {
            align: HorizontalPositionAlign.CENTER,
        },
        verticalPosition: {
            align: VerticalPositionAlign.CENTER,
        },
        allowOverlap: true,
        behindDocument: true,
    },
});

function makeRowTemplate1(title: string, subTitle: string = '', height: number = 0): TableRow {
    let row = new TableRow({
        height: {
            height: height !== 0 ? height : 480,
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
                columnSpan: 3,
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

function makeRowTemplate2(titles: string[]): TableRow {
    let tableCells: TableCell[] = [];
    for (let i = 0; i < 4; i++) {
        tableCells.push(
            new TableCell({
                verticalAlign: VerticalAlign.CENTER,
                width: {
                    size: i === 2 ? 2565 : 2115,
                    type: WidthType.DXA,
                },
                children: [
                    new Paragraph({
                        text: titles[i],
                        alignment: AlignmentType.CENTER,
                    }),
                ],
            }),
        );
    }

    let row = new TableRow({
        height: {
            height: 480,
            rule: HeightRule.EXACT,
        },
        children: tableCells,
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
                    verticalAlign: VerticalAlign.CENTER,
                    columnSpan: 4,
                    children: [
                        new Paragraph({
                            children: [
                                new TextRun({
                                    text: '라 이 선 스 내 역',
                                    bold: true,
                                    size: 32,
                                }),
                            ],
                            alignment: AlignmentType.CENTER,
                        }),
                    ],
                }),
            ],
        }),
        makeRowTemplate2(['항  목', 'Version', '신청수량', '비       고']),
        makeRowTemplate2(['마스터', '', '', '이중화는 별도 기재']),
        makeRowTemplate2(['저장/분배', '', '', '이중화는 별도 기재']),
        makeRowTemplate2(['클라이언트', '', '', '']),
        makeRowTemplate2(['카메라', '', '', '']),
        makeRowTemplate2(['SDK', '', '', '']),
        makeRowTemplate2(['TSMV', '', '', '']),
        makeRowTemplate2(['', '', '', '']),
        makeRowTemplate2(['', '', '', '']),
        makeRowTemplate2(['', '', '', '']),
        makeRowTemplate1('발주 요청 사항', '', 1440),
    ],
});

doc.addSection({
    children: [
        new Paragraph(background),
        new Paragraph({
            children: [
                new TextRun({
                    text: '라이선스 신청서',
                    underline: {
                        type: UnderlineType.SINGLE,
                    },
                    bold: true,
                    size: 52,
                }),
            ],
            alignment: AlignmentType.CENTER,
        }),
        new Paragraph({
            text: '',
            spacing: {
                before: 300,
            },
        }),
        table,
        new Paragraph({
            text: '',
            spacing: {
                before: 300,
            },
        }),
        new Paragraph({
            text: '2021년 01월 07일',
            heading: HeadingLevel.HEADING_3,
            spacing: {
                before: 300,
            },
            alignment: AlignmentType.CENTER,
        }),
        new Paragraph({
            children: [
                new TextRun({
                    text: '업체명 :            (인)',
                }),
                new TextRun({
                    text: '담당자 :                 ',
                }),
            ],
            alignment: AlignmentType.CENTER,
        }),
        // new Paragraph({
        //     text: '담당자 :                 ',
        //     heading: HeadingLevel.HEADING_1,
        //     alignment: AlignmentType.CENTER,
        //     spacing: {
        //         before: 300,
        //     },
        // }),
        new Paragraph(logo),
    ],
});

Packer.toBuffer(doc).then(buffer => {
    fs.writeFileSync('MyDocument.docx', buffer);
});
