//https://www.npmjs.com/package/docx

import {
    Alignment,
    AlignmentType,
    Border,
    BorderStyle,
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

const noBorderStyle = {
    top: {
        style: BorderStyle.NONE,
        size: 0,
        color: 'black',
    },
    bottom: {
        style: BorderStyle.NONE,
        size: 0,
        color: 'black',
    },
    left: {
        style: BorderStyle.NONE,
        size: 0,
        color: 'black',
    },
    right: {
        style: BorderStyle.NONE,
        size: 0,
        color: 'black',
    },
};

function makeRowTemplate1(title: string, subTitle: string = '', height: number): TableRow {
    let cellWidth: number = 15;
    let fontSize: number = 32;
    let row = new TableRow({
        height: {
            height: height,
            rule: HeightRule.EXACT,
        },
        children: [
            new TableCell({
                width: {
                    size: cellWidth * 70,
                    type: WidthType.DXA,
                },
                children: [],
                borders: noBorderStyle,
            }),
            new TableCell({
                verticalAlign: VerticalAlign.CENTER,
                width: {
                    size: cellWidth * 100,
                    type: WidthType.DXA,
                },
                children: [
                    new Paragraph({
                        children: [
                            new TextRun({
                                text: title,
                                size: fontSize,
                            }),
                        ],
                        alignment: AlignmentType.CENTER,
                    }),
                ],
                borders: noBorderStyle,
            }),
            new TableCell({
                columnSpan: 3,
                verticalAlign: VerticalAlign.CENTER,
                width: {
                    size: cellWidth * 400,
                    type: WidthType.DXA,
                },
                children: [
                    new Paragraph({
                        children: [
                            new TextRun({
                                text: subTitle,
                                size: fontSize,
                            }),
                        ],
                        alignment: AlignmentType.LEFT,
                    }),
                ],
                borders: noBorderStyle,
            }),
        ],
    });

    return row;
}

function makeRowTemplate2(title: string, height: number = 0): TableRow {
    let cellWidth: number = 15 * 1000;
    let fontSize: number = 32;
    let row = new TableRow({
        height: {
            height: height !== 0 ? height : 480,
            rule: HeightRule.EXACT,
        },
        children: [
            new TableCell({
                verticalAlign: VerticalAlign.CENTER,
                width: {
                    size: cellWidth,
                    type: WidthType.DXA,
                },
                children: [
                    new Paragraph({
                        children: [
                            new TextRun({
                                text: title,
                                size: fontSize,
                                bold: true,
                            }),
                        ],
                        alignment: AlignmentType.CENTER,
                    }),
                ],
                borders: noBorderStyle,
            }),
        ],
    });

    return row;
}

function makeRowTemplate3(titles: string[], height: number = 0): TableRow {
    let fontSize: number = 28;

    let tableCells: TableCell[] = [
        new TableCell({
            width: {
                size: 15 * 70,
                type: WidthType.DXA,
            },
            children: [],
            borders: noBorderStyle,
        }),
    ];
    let cnt: number = 0;
    for (let title of titles) {
        let cellWidth: number = 15 * 120;
        switch (cnt) {
            case 0: {
                cellWidth = 15 * 120;
                break;
            }
            case 1: {
                cellWidth = 15 * 240;
                break;
            }
            case 2: {
                cellWidth = 15 * 70;
                break;
            }
            case 3: {
                cellWidth = 15 * 40;
                break;
            }
        }
        tableCells.push(
            new TableCell({
                verticalAlign: VerticalAlign.CENTER,
                width: {
                    size: cellWidth,
                    type: WidthType.DXA,
                },
                children: [
                    new Paragraph({
                        children: [
                            new TextRun({
                                text: title,
                                size: fontSize,
                            }),
                        ],
                        alignment: AlignmentType.LEFT,
                    }),
                ],
                borders: noBorderStyle,
            }),
        );
        cnt++;
    }
    let row = new TableRow({
        height: {
            height: height !== 0 ? height : 620,
            rule: HeightRule.EXACT,
        },
        children: tableCells,
    });

    return row;
}

export function saveLicenseCertificate() {
    const doc = new Document();

    const logo = Media.addImage(doc, fs.readFileSync('./resources/innodep_logo2.jpg'), 197, 40, {
        floating: {
            horizontalPosition: {
                offset: 201440 * 6,
            },
            verticalPosition: {
                offset: 201440 * 45,
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

    const innodepStamp = Media.addImage(doc, fs.readFileSync('./resources/innodep_stamp.jpg'), 90, 90, {
        floating: {
            horizontalPosition: {
                offset: 201440 * 24,
            },
            verticalPosition: {
                offset: 201440 * 39,
            },
            wrap: {
                type: TextWrappingType.NONE,
                side: TextWrappingSide.BOTH_SIDES,
            },
        },
    });

    const vurixBg = Media.addImage(doc, fs.readFileSync('./resources/vurix_bg.png'), 440, 140, {
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

    const background = Media.addImage(doc, fs.readFileSync('./resources/frame2.png'), 750, 1080, {
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

    const mainTable = new Table({
        rows: [
            makeRowTemplate1('인증번호 :', ' IDLC200624-01', 520),
            makeRowTemplate1('발급일자 :', ' 2020년 6월 24일', 520),
            makeRowTemplate1('고 객 명 :', ' 전라북도 김제시', 520),
            makeRowTemplate1('사 업 명 :', ' 2020년 불법쓰레기 투기방지 CCTV 설치공사 관급자재 구입', 520 * 3),
        ],
    });

    const titleTable = new Table({
        rows: [makeRowTemplate2('─　　라 이 선 스　　내 역　　─')],
    });

    //INFO: 5개 고정
    const detailTable = new Table({
        rows: [
            makeRowTemplate3(['카메라', 'VURIX-ENT-DCAC', '8', 'LC']),
            makeRowTemplate3(['', '', '', '']),
            makeRowTemplate3(['', '', '', '']),
            makeRowTemplate3(['', '', '', '']),
            makeRowTemplate3(['', '', '', '']),
        ],
    });

    doc.addSection({
        children: [
            new Paragraph(background),
            new Paragraph(vurixBg),
            new Paragraph({
                children: [
                    new TextRun({
                        text: '라이선스 인증서',
                        underline: {
                            type: UnderlineType.DOUBLE,
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
            mainTable,
            new Paragraph({
                text: '',
                spacing: {
                    before: 600,
                },
            }),
            titleTable,
            new Paragraph({
                text: '',
                spacing: {
                    before: 300,
                },
            }),
            detailTable,
            new Paragraph({
                text: '',
                spacing: {
                    before: 600,
                },
            }),
            new Paragraph(logo),
            new Paragraph(innodepStamp),
            new Paragraph({
                children: [
                    new TextRun({
                        text: '이 노 뎁　　주식회사',
                        size: 30,
                    }),
                ],
                alignment: AlignmentType.CENTER,
            }),
            new Paragraph({
                children: [
                    new TextRun({
                        text: '대표 이사　 이 성 진',
                        size: 30,
                    }),
                ],
                alignment: AlignmentType.CENTER,
            }),
            new Paragraph({
                text: '',
                spacing: {
                    before: 400,
                },
            }),
            new Paragraph({
                children: [
                    new TextRun({
                        text: '이  노  뎁   주  식  회  사',
                        size: 30,
                    }),
                ],
                alignment: AlignmentType.RIGHT,
            }),
            new Paragraph({
                children: [
                    new TextRun({
                        text: '서울특별시 구로구 디지털로31길 61, 드림마크원데이터센타 5층 Tel. 02)2109-6866  Fax. 02)6336-3368',
                        size: 20,
                        bold: true,
                    }),
                ],
                alignment: AlignmentType.CENTER,
            }),
        ],
    });

    Packer.toBuffer(doc).then(buffer => {
        fs.writeFileSync('LicenseCertificate.docx', buffer);
    });
}
