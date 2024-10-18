//
//  LxwError.swift
//
//
//  Created by Yoshinori Takada on 2024/08/29.
//

import libxlsxwriter

enum LxwError: Error {
    case noError
    case memoryMallocFailed
    case creatingXlsxFile
    case creatingTmpfile
    case readingTmpfile
    case zipFileOperation
    case zipParameterError
    case zipBadZipFile
    case zipInternalError
    case zipFileAdd
    case zipClose
    case featureNotSupported
    case nullParameterIgnored
    case parameterValidation
    case parameterIsEmpty
    case sheetnameLengthExceeded
    case invalidSheetNameCharacter
    case sheetnameStartEndApostrophe
    case sheetnameAlreadyUsed
    case _32StringLengthExceeded
    case _128StringLengthExceeded
    case _255StringLengthExceeded
    case maxStringLengthExceeded
    case sharedStringIndexNotFound
    case worksheetIndexOutOfRange
    case worksheetMaxUrlLengthExceeded
    case worksheetMaxNumberUrlsExceeded
    case imageDimensions
}
