const fs = require("fs");
const express = require('express');
const { Document, 
        Packer, 
        Paragraph, 
        TextRun, 
        Tab, 
        Table, 
        TableRow, 
        TableCell,
        HeightRule, 
        HeadingLevel, 
        AlignmentType,
        WidthType, 
        TabStopType, 
        TabStopPosition } = require("docx");
        

const app = express();

app.use(express.json())

app.get('/', (req, res) => {
    res.send('Hello World')
})

app.post('/convert-to-docx', (req, res) => {
    try {
        const { data } = req.body;
        const { projectRows } = data;

        //Document Children Initialisation
        const children = [
            new Paragraph({
                children: [
                    new TextRun({
                        text: "Hot Sheet",
                        color: "#002F8C",
                        bold: true,
                        heading: HeadingLevel.HEADING_1,
                        size: 25,
                    }),
                ],
                alignment: AlignmentType.CENTER,
            }),
        ];

        if (projectRows?.length) {
            projectRows.map((rowData, index) => {
                const { defaultFields, customFields } = rowData;

                // Appending Default Fields
                const defaultFieldsList = Object.keys(defaultFields)
                if (defaultFieldsList?.length) {
                    defaultFieldsList.map((field, index1) => {
                        children.push(new Paragraph({
                            children: [
                                new TextRun({
                                    text: index1 == 0 ? index + 1 + ')' : '',
                                    size: 16,
                                    bold: true,
                                }),
                                new TextRun({
                                    text: "\t" + field.charAt(0).toUpperCase() + field.slice(1) + ':',
                                    size: 16,
                                    bold: true
                                }),
                                new TextRun({
                                    children: [new Tab(), "\t" + defaultFields[field]],
                                    size: 16
                                }),
                            ],
                            tabStops: [
                                {
                                    type: TabStopType.LEFT,
                                    position: TabStopPosition.LEFT,
                                },
                                {
                                    type: TabStopType.LEFT,
                                    position: 500,
                                },
                                {
                                    type: TabStopType.LEFT,
                                    position: 2000,
                                },
                            ],
                        }))
                    })
                }

                // Adding customFields
                const customFieldsList = Object.keys(customFields)
                if (customFieldsList?.length) {
                    children.push(new Paragraph({
                        children: [
                            new TextRun({
                                text: '',
                                size: 16,
                                bold: true,
                                break: 1
                            }),
                            new TextRun({
                                text: "\t" + 'CustomFields:',
                                size: 16,
                                bold: true
                            }),
                            new TextRun({
                                text: '',
                            }),
                        ],
                        tabStops: [
                            {
                                type: TabStopType.LEFT,
                                position: TabStopPosition.LEFT,
                            },
                            {
                                type: TabStopType.LEFT,
                                position: 500,
                            },
                        ],
                    }))

                    //Creating table
                    const rows = [
                        new TableRow({
                            children: [
                                new TableCell({
                                    children: [new Paragraph({
                                        children: [
                                            new TextRun({
                                                text: "\tName",
                                                bold: true,
                                                size: 16,
                                                allCaps: true
                                            })
                                        ]
                                    })],
                                }),
                                new TableCell({
                                    children: [new Paragraph({
                                        children: [
                                            new TextRun({
                                                text: "\tValue",
                                                bold: true,
                                                size: 16,
                                                allCaps: true
                                            })
                                        ]
                                    })],
                                }),
                            ],
                            tableHeader: true,
                        })
                    ]

                    //Adding Rows to the Table
                    customFieldsList.map((rowData) => {
                        rows.push(new TableRow({
                            children: [
                                new TableCell({
                                    children: [new Paragraph({
                                        children: [
                                            new TextRun({
                                                text: rowData || '',
                                                size: 16
                                            })
                                        ]
                                    })],
                                }),
                                new TableCell({
                                    children: [new Paragraph({
                                        children: [
                                            new TextRun({
                                                text: customFields[rowData] || '',
                                                size: 16
                                            })
                                        ]
                                    })],
                                }),
                            ],
                        }))
                    })

                    //Creating Table
                    children.push(new Table({
                        rows,
                        width: {
                            size: 4535,
                            type: WidthType.DXA,
                        },
                        indent: {
                            size: 500,
                            type: WidthType.DXA,
                        },
                        height: {
                            value: 1000,
                            rule: HeightRule.ATLEAST
                        }
                    }))
                }

                // Adding Break
                if (defaultFieldsList.length || customFieldsList.length) {
                    children.push(new Paragraph({
                        children: [
                            new TextRun({
                                text: '',
                                break: 1
                            }),
                        ],
                    }))
                }
            })
        }

        //Creating Document
        const doc = new Document({
            sections: [
                {
                    properties: {},
                    children
                },
            ],
        });

        // Test to Save the File in Local
        // Pack the document into a buffer
        // Packer.toBuffer(doc).then((buffer) => {
        //     // Write the buffer to a .docx file
        //     fs.writeFileSync("document.docx", buffer);
        //     console.log("Document created successfully!");
        // }).catch((err) => {
        //     console.error("Error:", err);
        // });

        // Used to export the file into a .docx file
        Packer.toBuffer(doc).then((buffer) => {
            res.status(200).json({
                status: "Success",
                message: "Converted Json to Docx Format",
                buffer,
            })
        });
    } catch (error) {
        res.status(400).json({
            status: "Failed",
            message: "Conversion Json to Docx Format Failed " + error
        })
    }
})

app.listen(8000, () => {
    console.log('JsonToDocx is listening on 8000')
})