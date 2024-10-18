// swift-tools-version:5.2
import PackageDescription

let package = Package(
    name: "xlsxwriter.swift",
    products: [
        .library(name: "xlsxwriter", targets: ["xlsxwriter"])
    ],
    dependencies: [
        .package(url: "https://github.com/jmcnamara/libxlsxwriter", .branchItem("main")),
        .package(url: "https://github.com/apple/swift-docc-plugin", .upToNextMajor(from: "1.0.0")),
        .package(url: "https://github.com/CoreOffice/CoreXLSX.git", .upToNextMajor(from: "0.14.1")),
    ],
    targets: [
        .target(
            name: "xlsxwriter",
            dependencies: ["libxlsxwriter"],
            cSettings: [(.define("_CRT_SECURE_NO_WARNINGS", .when(platforms: [.windows])))]
        ),
        .testTarget(
            name: "xlsxwriterTests",
            dependencies: ["xlsxwriter", "CoreXLSX"]
        ),
    ]
)
