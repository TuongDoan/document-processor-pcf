import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as React from "react";
import {
  FilesImportControl,
  IFilesImportControlProps,
} from "./FilesImportControl";
import { Theme, webLightTheme } from "@fluentui/react-components";
import { parseExcelDynamic, parseExcelFixed } from "./Excel";

export class DocumentProcessor
  implements ComponentFramework.ReactControl<IInputs, IOutputs>
{
  private notifyOutputChanged!: () => void;
  private filesAsJSON: string | null = null;
  private excelOutput: string | null = null;
  private excelOutputs: string[] | null = null;

  private rangeMode: boolean;
  private targetRange: string;
  private autoHeader: boolean;

  public init(
    _context: ComponentFramework.Context<IInputs>,
    notifyOutputChanged: () => void,
    _state: ComponentFramework.Dictionary
  ): void {
    this.notifyOutputChanged = notifyOutputChanged;
  }

  public updateView(
    context: ComponentFramework.Context<IInputs>
  ): React.ReactElement {
    const theme: Theme =
      (context.fluentDesignLanguage?.tokenTheme as Theme | undefined) ??
      webLightTheme;

    this.rangeMode =
      (context.parameters.RangeMode?.raw as unknown as boolean) ?? false;
    this.autoHeader =
      (context.parameters.AutoHeader?.raw as unknown as boolean) ?? false;
    this.targetRange = context.parameters.targetRange?.raw ?? "";

    const props: IFilesImportControlProps = {
      buttonText: context.parameters.Text?.raw ?? "File upload",
      buttonIcon: context.parameters.ButtonIcon?.raw ?? "Attach",
      buttonIconStyle: context.parameters.IconStyle?.raw ?? "Outline",
      buttonAppearance: context.parameters.Appearance?.raw ?? "Primary",
      buttonAlign: context.parameters.Align?.raw ?? "Center",
      buttonFontWeight: context.parameters.FontWeight?.raw ?? "Normal",
      buttonVisible: context.parameters.Visible?.raw ?? true,
      buttonisDisabled: context.mode.isControlDisabled,
      buttonWidth: context.parameters.Width?.raw ?? 192,
      buttonHeight: context.parameters.Height?.raw ?? 32,
      buttonShowSecondaryContent: context.parameters.ShowSecondaryContent?.raw ?? false,
      buttonSecondaryContent: context.parameters.SecondaryContent?.raw ?? "",
      buttonShowActionSpinner: context.parameters.ShowActionSpinner?.raw ?? true,
      buttonIconPosition: context.parameters.IconPosition?.raw ?? "Before",
      buttonShape: context.parameters.Shape?.raw ?? "Rounded",
      buttonButtonSize: context.parameters.ButtonSize?.raw ?? "Medium",
      canvasAppCurrentTheme: theme,
      buttonAllowMultipleFiles: context.parameters.AllowMultipleFiles?.raw ?? false,
      buttonAllowedFileTypes: context.parameters.AllowedFileTypes?.raw ?? "*",
      buttonAllowDropFiles: context.parameters.AllowDropFiles?.raw ?? false,
      buttonAllowDropFilesText: context.parameters.AllowDropFilesText?.raw ??
        "Drop file vào đây nè má!",
      onEvent: (e) => {
        void this.handleFileUpload(e);
      },
    };

    return React.createElement(FilesImportControl, props);
  }

  private handleFileUpload = async (event: {
    filesJSON: { name: string; binary: ArrayBuffer; base64: string }[];
  }): Promise<void> => {
    this.filesAsJSON = JSON.stringify(
      event.filesJSON.map((f) => ({ name: f.name, base64: f.base64 }))
    );

    try {
      const files = event.filesJSON;
      if (Array.isArray(files) && files.length > 0) {
        const range = this.targetRange;
        const autoHeader = this.autoHeader;

        const results: unknown[] = files.map((file) => {
          if (file?.name && file?.binary && file.binary.byteLength > 0) {
            if (this.rangeMode && range) {
              return parseExcelFixed(file.binary, range, file.name, autoHeader);
            } else {
              return parseExcelDynamic(file.binary, file.name);
            }
          } else {
            return "";
          }
        });
        this.excelOutputs = results as string[];
        this.excelOutput = JSON.stringify(results);
      } else {
        this.excelOutputs = [];
        this.excelOutput = "";
      }
    } catch (err) {
      this.excelOutputs = [];
      this.excelOutput = "";
    }
    this.notifyOutputChanged();
    return Promise.resolve();
  };

  public getOutputs(): IOutputs {
    let excelOutputParsed: unknown[] = [];
    if (this.excelOutput) {
      try {
        const parsed: unknown = JSON.parse(this.excelOutput);
        if (Array.isArray(parsed)) {
          excelOutputParsed = parsed;
        } else {
          excelOutputParsed = [parsed];
        }
      } catch {
        excelOutputParsed = [this.excelOutput];
      }
    }
    return {
      FilesAsJSON: this.filesAsJSON ?? "",
      ExcelOutput: JSON.stringify(excelOutputParsed),
    };
  }

  public destroy(): void {
    // Cleanup if necessary
    }
}
