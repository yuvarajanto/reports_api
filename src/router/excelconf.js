const express = require('express');
const quote = require('../models/quotesmodel');
const newquote = require('../models/newquotemodel');
const quotecc = require('../models/quoteccmodel');
const ExcelJS = require('exceljs');


function excelData(data, product) {
    const excelFormat = data.map(item => {
        const counts = {
            "Draft": '',
            "DRAFT": '',
            "Order Submitted": '',
            "Order Placed": '',
            "Order Completed": '',
            "Order Implemented": ''
        };

        item.statusCounts.forEach(statusCount => {
            counts[statusCount.status] = statusCount.count;
        });

        return {
            "Date": item.createdDate,
            "Product": product,
            "No. of Orders": item.totalCount,
            "DRAFT": counts["DRAFT"] || counts["Draft"],
            "Order Submitted": counts["Order Submitted"],
            "Order Completed": counts["Order Completed"],
            "Order Placed": counts["Order Placed"],
            "Order Implemented": counts["Order Implemented"],
            "OBValue(ARC)": item.obvalueArc,
            "OBValue(OTC)": item.obvalueOtc
        };
    });
    return excelFormat;
}

const generateExcel = async () => {
    try {
        const currentDate = new Date();
        const currentYear = currentDate.getFullYear();
        const currentMonth = currentDate.getMonth() + 1; 

        const currentQuarter = Math.ceil(currentMonth / 3);
        const quoteData = await quote.aggregate([
            {
                $addFields: {
                  year: { $year: "$createdDate" },
                  quarter: {
                    $switch: {
                      branches: [
                        { case: { $lte: [{ $month: "$createdDate" }, 3] }, then: 1 },
                        { case: { $lte: [{ $month: "$createdDate" }, 6] }, then: 2 },
                        { case: { $lte: [{ $month: "$createdDate" }, 9] }, then: 3 },
                        { case: { $lte: [{ $month: "$createdDate" }, 12] }, then: 4 }
                      ],
                      default: null
                    }
                  }
                }
              },
              {
                // Stage 2: Match documents for the current year and quarter
                $match: {
                  year: currentYear,
                  quarter: currentQuarter
                }
              },
            {
                $group: {
                    _id: {
                        createdDate: { $dateToString: { format: "%d %B %Y", date: "$createdDate" } },
                        status: '$status'
                    },
                    count: { $sum: 1 },
                    totalArc: { $sum: { $cond: { if: { $or: [{ $eq: ["$status", "Order Completed"] }, { $eq: ["$status", "Order Implemented"] }, { $eq: ["$status", "Order placed"] }] }, then: "$poDetails.arc", else: 0 } } },
                    totalOtc: { $sum: { $cond: { if: { $or: [{ $eq: ["$status", "Order Completed"] }, { $eq: ["$status", "Order Implemented"] }] }, then: "$poDetails.otc", else: 0 } } }
                }
            },
            {
                $group: {
                    _id: '$_id.createdDate',
                    statusCounts: {
                        $push: {
                            status: '$_id.status',
                            count: '$count'
                        }
                    },
                    totalCount: { $sum: { $cond: { if: { $or: [{ $eq: ["$_id.status", "DRAFT"] }, { $eq: ["$_id.status", "Order Completed"] }, { $eq: ["$_id.status", "Order Placed"] }, { $eq: ["$_id.status", "Order Implemented"] }] }, then: '$count', else: 0 } } },
                    obvalueArc: { $sum: '$totalArc' },
                    obvalueOtc: { $sum: '$totalOtc' }
                }
            },
            {
                $sort: {
                    "_id": 1
                }
            },
            {
                $project: {
                    _id: 0,
                    createdDate: '$_id',
                    statusCounts: 1,
                    totalCount: 1,
                    obvalueArc: 1,
                    obvalueOtc: 1
                }
            }
        ]);
        const newquoteData = await newquote.aggregate([
            {
                $addFields: {
                  createdDate: {
                    $cond: {
                        if: { $eq: [{ $type: "$createdDate" }, "string"] }, 
                      then: { $dateFromString: { dateString: "$createdDate" } }, 
                      else: "$createdDate" 
                    }
                  }
                }
              },
            {
                $addFields: {
                  year: { $year: "$createdDate" },
                  quarter: {
                    $switch: {
                      branches: [
                        { case: { $lte: [{ $month: "$createdDate" }, 3] }, then: 1 },
                        { case: { $lte: [{ $month: "$createdDate" }, 6] }, then: 2 },
                        { case: { $lte: [{ $month: "$createdDate" }, 9] }, then: 3 },
                        { case: { $lte: [{ $month: "$createdDate" }, 12] }, then: 4 }
                      ],
                      default: null
                    }
                  }
                }
              },
              {
                // Stage 2: Match documents for the current year and quarter
                $match: {
                  year: currentYear,
                  quarter: currentQuarter
                }
              },
            {
                $unwind: '$vas_link'
            },
            {
                $group: {
                    _id: {
                        createdDate: { $dateToString: { format: "%d %B %Y", date: { $toDate: "$createdDate" } } },
                        status: '$status'
                    },
                    count: { $sum: 1 },
                    totalArc: { $sum: { $cond: { if: { $or: [{ $eq: ["$status", "Order Completed"] }, { $eq: ["$status", "Order Implemented"] }, { $eq: ["$status", "Order Placed"] }, { $eq: ["$status", "Order Submitted"] }] }, then: { $add: [{ $toDouble: "$bandwidthPriceValue.totalArc" }, { $sum: "$vas_link.getVas.arc" }] }, else: 0 } } },
                    totalOtc: { $sum: { $cond: { if: { $or: [{ $eq: ["$status", "Order Completed"] }, { $eq: ["$status", "Order Implemented"] }, { $eq: ["$status", "Order Placed"] }, { $eq: ["$status", "Order Submitted"] }] }, then: { $add: [{ $toDouble: "$bandwidthPriceValue.totalOtc" }, { $sum: "$vas_link.getVas.otc" }] }, else: 0 } } }
                }
            },
            {
                $group: {
                    _id: '$_id.createdDate',
                    statusCounts: {
                        $push: {
                            status: '$_id.status',
                            count: '$count'
                        }
                    },
                    totalCount: { $sum: { $cond: { if: { $or: [{ $eq: ["$_id.status", "DRAFT"] }, { $eq: ["$_id.status", "Order Completed"] }, { $eq: ["$_id.status", "Order Placed"] }, { $eq: ["$_id.status", "Order Implemented"] }] }, then: '$count', else: 0 } } },
                    obvalueArc: { $sum: { $toDouble: '$totalArc' } },
                    obvalueOtc: { $sum: { $toDouble: '$totalOtc' } },
                }
            },
            {
                $project: {
                    _id: 0,
                    createdDate: '$_id',
                    statusCounts: 1,
                    totalCount: 1,
                    obvalueArc: 1,
                    obvalueOtc: 1
                }
            }
        ]);
        const quoteccData = await quotecc.aggregate([
            {
                $addFields: {
                  year: { $year: "$createdDate" },
                  quarter: {
                    $switch: {
                      branches: [
                        { case: { $lte: [{ $month: "$createdDate" }, 3] }, then: 1 },
                        { case: { $lte: [{ $month: "$createdDate" }, 6] }, then: 2 },
                        { case: { $lte: [{ $month: "$createdDate" }, 9] }, then: 3 },
                        { case: { $lte: [{ $month: "$createdDate" }, 12] }, then: 4 }
                      ],
                      default: null
                    }
                  }
                }
              },
              {
                // Stage 2: Match documents for the current year and quarter
                $match: {
                  year: currentYear,
                  quarter: currentQuarter
                }
              },
            {
                $group: {
                    _id: {
                        createdDate: { $dateToString: { format: "%d %B %Y", date: "$createdDate" } },
                        status: '$status'
                    },
                    count: { $sum: 1 },
                    totalArc: { $sum: { $cond: { if: { $or: [{ $eq: ["$status", "Order Completed"] }, { $eq: ["$status", "Order Implemented"] }, { $eq: ["$status", "Order Placed"] }, { $eq: ["$status", "Order Submitted"] }] }, then: "$arc", else: 0 } } },
                    totalOtc: { $sum: { $cond: { if: { $or: [{ $eq: ["$status", "Order Completed"] }, { $eq: ["$status", "Order Implemented"] }, { $eq: ["$status", "Order Placed"] }, { $eq: ["$status", "Order Submitted"] }] }, then: "$otc", else: 0 } } }
                }
            },
            {
                $group: {
                    _id: '$_id.createdDate',
                    statusCounts: {
                        $push: {
                            status: '$_id.status',
                            count: '$count'
                        }
                    },
                    totalCount: { $sum: { $cond: { if: { $or: [{ $eq: ["$_id.status", "Draft"] }, { $eq: ["$_id.status", "Order Completed"] }, { $eq: ["$_id.status", "Order Placed"] }, { $eq: ["$_id.status", "Order Implemented"] }] }, then: '$count', else: 0 } } },
                    obvalueArc: { $sum: '$totalArc' },
                    obvalueOtc: { $sum: '$totalOtc' }
                }
            },
            {
                $sort: {
                    "_id": 1
                }
            },
            {
                $project: {
                    _id: 0,
                    createdDate: '$_id',
                    statusCounts: 1,
                    totalCount: 1,
                    obvalueArc: 1,
                    obvalueOtc: 1
                }
            }
        ]);
        const quoteExcel = excelData(quoteData, "NSE Migrate P2P");
        const newquoteExcel = excelData(newquoteData, "NSE NEW");
        const quoteccExcel = excelData(quoteccData, "CrossConnect");
        console.log(JSON.stringify(quoteData, null, 2))
        const combinedData = [...quoteExcel, ...newquoteExcel, ...quoteccExcel];
        combinedData.sort((a, b) => new Date(a.Date) - new Date(b.Date));
        // console.log(JSON.stringify(combinedData,null,2));
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Summary');


        worksheet.columns = [
            { header: 'Date', key: 'Date', width: 20 },
            { header: 'Product', key: 'Product', width: 20 },
            { header: 'No. of Orders', key: 'NoOfOrders', width: 15 },
            { header: 'DRAFT', key: 'DRAFT', width: 10 },
            { header: 'Order Submitted', key: 'OrderSubmitted', width: 18 },
            { header: 'Order Completed', key: 'OrderCompleted', width: 18 },
            { header: 'Order Placed', key: 'OrderPlaced', width: 15 },
            { header: 'OBValue(ARC)', key: 'OBValueARC', width: 20 },
            { header: 'OBValue(OTC)', key: 'OBValueOTC', width: 20 }
        ];

        worksheet.getRow(1).eachCell((cell) => {
            cell.font = { bold: true, color: { argb: '000000' } };
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'D3D3D3' }
            };
            cell.alignment = { vertical: 'middle', horizontal: 'center' };
        });

        let prevDate = null;
        let mergeStartRow = null;

        combinedData.forEach((row, index) => {
            const currentRow = worksheet.addRow({
                Date: row.Date,
                Product: row.Product,
                NoOfOrders: row['No. of Orders'],
                DRAFT: row.DRAFT,
                OrderSubmitted: row['Order Submitted'],
                OrderCompleted: row['Order Completed'],
                OrderPlaced: row['Order Placed'],
                OBValueARC: row['OBValue(ARC)'] || '',
                OBValueOTC: row['OBValue(OTC)'] || ''
            });

            currentRow.eachCell((cell) => {
                cell.alignment = { vertical: 'middle', horizontal: 'center' };
                cell.border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' }
                };
            });

            const currentRowNumber = currentRow.number;

            if (prevDate === row.Date) {
            } else {

                if (mergeStartRow !== null && mergeStartRow !== currentRowNumber - 1) {
                    worksheet.mergeCells(`A${mergeStartRow}:A${currentRowNumber - 1}`);
                    // console.log(`Merging cells from A${mergeStartRow} to A${currentRowNumber - 1}`);
                }

                mergeStartRow = currentRowNumber;
            }
            prevDate = row.Date;
            if (index === combinedData.length - 1 && mergeStartRow !== currentRowNumber) {
                worksheet.mergeCells(`A${mergeStartRow}:A${currentRowNumber}`);
                // console.log(`Merging cells from A${mergeStartRow} to A${currentRowNumber}`);
            }
        });

        // res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        // res.setHeader('Content-Disposition', 'attachment; filename=actionRequired.xlsx');

        const buffer = await workbook.xlsx.writeBuffer();
        return buffer;

    } catch (err) {
        console.log(err);
    }
};


module.exports = { generateExcel};