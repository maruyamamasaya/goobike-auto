import {
  excelSheetHeaders,
  excelSheetInstructions,
  excelSheetRows,
  excelSheetSample,
} from "./fixtures/excelSheet";

describe("Excel sheet test data", () => {
  it("should contain three rows", () => {
    expect(excelSheetRows).toHaveLength(3);
  });

  it("should capture the expected header information", () => {
    expect(excelSheetHeaders[0]).toBe("No.");
    expect(excelSheetHeaders).toEqual(
      expect.arrayContaining([
        "車種",
        "車体番号",
        "画像登録URL",
        "投稿結果",
        "記事URL",
      ])
    );
  });

  it("should capture the instruction row", () => {
    expect(excelSheetInstructions[1]).toBe("🔻を選択");
    expect(excelSheetInstructions).toContain("クリックして画像アップロード　※パスあり");
    expect(excelSheetInstructions).toContain("自動登録のため記入不要");
  });

  it("should capture the sample vehicle row", () => {
    expect(excelSheetSample).toContain("VTR250");
    expect(excelSheetSample).toContain("MC33-1303943");
    expect(excelSheetSample).toContain("https://57.180.200.150.sslip.io/uploader/01015");
    expect(excelSheetSample).toContain("TRUE");
    expect(excelSheetSample).toContain("FALSE");
  });
});
