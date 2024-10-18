//
//  PaperType.swift
//
//
//  Created by Yoshinori Takada on 2024/08/29.
//

import Foundation

/// The Excel paper format type.
public enum PaperType: UInt8 {
    case printerDefault = 0
    case letter
    case letterSmall
    case tabloid
    case ledger
    case legal
    case statement
    case executive
    case a3
    case a4
    case a4Small
    case a5
    case b4
    case b5
    case folio
    case quarto
    case unnamed1
    case unnamed2
    case note
    case envelope9
    case envelope10
    case envelope11
    case envelope12
    case envelope14
    case cSizeSheet
    case dSizeSheet
    case eSizeSheet
    case envelopeDL
    case envelopeC3
    case envelopeC4
    case envelopeC5
    case envelopeC6
    case envelopeC65
    case envelopeB4
    case envelopeB5
    case envelopeB6
    case envelope
    case monarch
    case envelopeInch
    case fanfold
    case germanStdFanfold
    case germanLegalFanfold
}
