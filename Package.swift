// swift-tools-version:5.2
import PackageDescription

let package = Package(
  name: "xlsxwriter.swift",
  products: [
    .library(name: "xlsxwriter", targets: ["xlsxwriter"])
  ],
  dependencies: [
    .package(url: "https://github.com/jmcnamara/libxlsxwriter", .branchItem("main")),
    .package(url: "https://github.com/apple/swift-docc-plugin", .exactItem("1.4.2")),
  ],
  targets: [
    .target(
      name: "xlsxwriter",
      dependencies: ["libxlsxwriter"],
      cSettings: [(.define("_CRT_SECURE_NO_WARNINGS", .when(platforms: [.windows])))]
    )
  ]
)
