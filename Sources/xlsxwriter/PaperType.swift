//
//  PaperType.swift
//
//
//  Created by Yoshinori Takada on 2024/08/29.
//

import Foundation

/// The Excel paper format type.
public enum PaperType: UInt8 {
  case PrinterDefault = 0  // Printer default
  case Letter  // 8 1/2 x 11 in
  case LetterSmall  // 8 1/2 x 11 in
  case Tabloid  // 11 x 17 in
  case Ledger  // 17 x 11 in
  case Legal  // 8 1/2 x 14 in
  case Statement  // 5 1/2 x 8 1/2 in
  case Executive  // 7 1/4 x 10 1/2 in
  case A3  // 297 x 420 mm
  case A4  // 210 x 297 mm
  case A4Small  // 210 x 297 mm
  case A5  // 148 x 210 mm
  case B4  // 250 x 354 mm
  case B5  // 182 x 257 mm
  case Folio  // 8 1/2 x 13 in
  case Quarto  // 215 x 275 mm
  case unnamed1  // 10x14 in
  case unnamed2  // 11x17 in
  case Note  // 8 1/2 x 11 in
  case Envelope_9  // 3 7/8 x 8 7/8
  case Envelope_10  // 4 1/8 x 9 1/2
  case Envelope_11  // 4 1/2 x 10 3/8
  case Envelope_12  // 4 3/4 x 11
  case Envelope_14  // 5 x 11 1/2
  case C_size_sheet  // ---
  case D_size_sheet  // ---
  case E_size_sheet  // ---
  case Envelope_DL  // 110 x 220 mm
  case Envelope_C3  // 324 x 458 mm
  case Envelope_C4  // 229 x 324 mm
  case Envelope_C5  // 162 x 229 mm
  case Envelope_C6  // 114 x 162 mm
  case Envelope_C65  // 114 x 229 mm
  case Envelope_B4  // 250 x 353 mm
  case Envelope_B5  // 176 x 250 mm
  case Envelope_B6  // 176 x 125 mm
  case Envelope  // 110 x 230 mm
  case Monarch  // 3.875 x 7.5 in
  case EnvelopeInch  // 3 5/8 x 6 1/2 in
  case Fanfold  // 14 7/8 x 11 in
  case German_Std_Fanfold  // 8 1/2 x 12 in
  case German_Legal_Fanfold  // 8 1/2 x 13 in
}
