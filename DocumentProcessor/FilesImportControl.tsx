import * as React from "react";
import { useState, createRef } from "react";
import {
  Caption1,
  Button,
  CompoundButton,
  Spinner,
  FluentProvider,
  Theme,
  webLightTheme,
} from "@fluentui/react-components";
import {
  AttachRegular,
  AttachFilled,
  CheckmarkFilled,
} from "@fluentui/react-icons";
import { iconRegularMapping, iconFilledMapping } from "./iconsMapping";
import { ButtonLoadingStateEnum } from "./utils";
import {
  getButtonAppearance,
  getButtonIconPosition,
  getButtonShape,
  getButtonSize,
  getButtonStyle,
  getDisplayMode,
} from "./utils";

export interface FilesImportEvent {
  filesJSON: { name: string; binary: ArrayBuffer; base64: string }[];
}

export interface IFilesImportControlProps {
  buttonText: string;
  buttonIcon: string;
  buttonIconStyle: string;
  buttonAppearance: string;
  buttonAlign: string;
  buttonVisible: boolean;
  buttonFontWeight: string;
  buttonisDisabled: boolean;
  buttonWidth: number;
  buttonHeight: number;
  buttonShowSecondaryContent: boolean;
  buttonSecondaryContent: string;
  buttonShowActionSpinner: boolean;
  buttonIconPosition: string;
  buttonShape: string;
  buttonButtonSize: string;
  buttonAllowMultipleFiles: boolean;
  buttonAllowedFileTypes: string;
  buttonAllowDropFiles: boolean;
  buttonAllowDropFilesText: string;
  canvasAppCurrentTheme: Theme;
  onEvent: (event: FilesImportEvent) => void;
}

export const FilesImportControl: React.FC<IFilesImportControlProps> = ({
  buttonText,
  buttonIcon,
  buttonIconStyle,
  buttonAppearance,
  buttonAlign,
  buttonVisible,
  buttonFontWeight,
  buttonisDisabled,
  buttonWidth,
  buttonHeight,
  buttonShowSecondaryContent,
  buttonSecondaryContent,
  buttonShowActionSpinner,
  buttonIconPosition,
  buttonShape,
  buttonButtonSize,
  buttonAllowMultipleFiles,
  buttonAllowedFileTypes,
  buttonAllowDropFiles,
  buttonAllowDropFilesText,
  canvasAppCurrentTheme,
  onEvent,
}) => {
  // THEME
  const _theme = canvasAppCurrentTheme?.fontFamilyBase?.trim()
    ? canvasAppCurrentTheme
    : webLightTheme;

  // BUTTON
  const [buttonLoadingState, setButtonLoadingState] =
    useState<ButtonLoadingStateEnum>(ButtonLoadingStateEnum.Initial);

  const getButtonIcon = () => {
    switch (buttonLoadingState) {
      case ButtonLoadingStateEnum.Loading:
        return <Spinner size="tiny" />;
      case ButtonLoadingStateEnum.Loaded:
        return <CheckmarkFilled />;
      case ButtonLoadingStateEnum.Initial:
      default:
        return buttonIconStyle === "1"
          ? iconFilledMapping[buttonIcon] ?? <AttachFilled />
          : iconRegularMapping[buttonIcon] ?? <AttachRegular />;
    }
  };

  const _buttonProps = {
    icon: getButtonIcon(),
    disabledFocusable: false, //buttonDisabledFocusable,
    appearance: getButtonAppearance(buttonAppearance),
    iconPosition: getButtonIconPosition(buttonIconPosition),
    shape: getButtonShape(buttonShape),
    size: getButtonSize(buttonButtonSize),
    style: getButtonStyle(
      buttonWidth,
      buttonHeight,
      buttonAlign,
      buttonFontWeight
    ),
    disabled: getDisplayMode(buttonLoadingState, buttonisDisabled),
  };

  // FILES

  const importFileRef = createRef<HTMLInputElement>();
  const readFile = (
    file: File
  ): Promise<{ name: string; size: number; content: string | ArrayBuffer }> => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onerror = () =>
        reject(new Error(`Error reading file: ${file.name}`));
      reader.onload = () => {
        // FileReader.result can be string | ArrayBuffer | null, but we only want string | ArrayBuffer
        if (reader.result === null) {
          reject(new Error(`FileReader result is null for file: ${file.name}`));
          return;
        }
        resolve({
          name: file.name,
          size: file.size,
          content: reader.result,
        });
      };
      if (file.name.endsWith(".xlsx")) {
        reader.readAsArrayBuffer(file);
      } else {
        reader.readAsDataURL(file);
      }
    });
  };

  const processFiles = async (files: FileList | File[]) => {
    if (!files || files.length === 0) return;

    try {
      if (buttonShowActionSpinner) {
        setButtonLoadingState(ButtonLoadingStateEnum.Loading);
      }

      const filesArray = await Promise.all(
        Array.from(files).map(async (file) => await readFile(file))
      );



      const filesJSON = filesArray.map((f) => {
        let base64 = "";
        let binary: ArrayBuffer = new ArrayBuffer(0);
        if (f.content instanceof ArrayBuffer) {
          binary = f.content;
          // Convert ArrayBuffer to base64
          let bin = "";
          const bytes = new Uint8Array(f.content);
          for (let i = 0; i < bytes.byteLength; i++) {
            bin += String.fromCharCode(bytes[i]);
          }
          base64 = btoa(bin);
        } else if (typeof f.content === "string") {

          const regex = /^data:.*;base64,(.*)$/;
          const match = regex.exec(f.content);
          base64 = match ? match[1] : f.content;
        }
        return {
          name: f.name,
          binary,
          base64,
        };
      });

      onEvent({ filesJSON });

      if (buttonShowActionSpinner) {
        setButtonLoadingState(ButtonLoadingStateEnum.Loaded);
      }
    } catch (error) {
      console.error("Error processing files:", error);
      setButtonLoadingState(ButtonLoadingStateEnum.Initial);
    }
  };

  const onFileChange = async (event: React.ChangeEvent<HTMLInputElement>) => {
    if (event.target.files) {
      await processFiles(event.target.files);
      event.target.value = "";
    }
  };

  const handleButtonClick = () => {
    setButtonLoadingState(ButtonLoadingStateEnum.Initial);
    importFileRef.current!.click();
  };

  // DRAG & DROP HANDLING
  const [isDragging, setIsDragging] = React.useState<boolean>(false);
  const dropZoneRef = React.useRef<HTMLDivElement>(null);

  const onDragEnter = (event: React.DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    setIsDragging(true);
  };

  const onDragOver = (event: React.DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    event.dataTransfer.dropEffect = "copy";
  };

  const onDragLeave = (event: React.DragEvent<HTMLDivElement>) => {
    if (!dropZoneRef.current!.contains(event.relatedTarget as Node)) {
      setIsDragging(false);
    }
  };

  const onDrop = async (event: React.DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    setIsDragging(false);

    if (event.dataTransfer.files) {
      await processFiles(event.dataTransfer.files);
    }
  };

  return (
    <div
      ref={dropZoneRef}
      onDragEnter={onDragEnter}
      onDragOver={onDragOver}
      onDragLeave={onDragLeave}
      onDrop={(e) => {
        void onDrop(e);
      }}
      style={{
        border:
          isDragging && buttonAllowDropFiles
            ? `1px dashed ${_theme.colorBrandBackground}`
            : "0px none transparent",
        padding: isDragging && buttonAllowDropFiles ? `10px` : "0px",
        borderRadius: isDragging && buttonAllowDropFiles ? `5px` : "0px",
        height:
          isDragging && buttonAllowDropFiles ? buttonHeight + 10 : buttonHeight,
        width:
          isDragging && buttonAllowDropFiles ? buttonWidth + 10 : buttonWidth,
        textAlign: "center",
        transition: "border 1s ease-in-out",
        backgroundColor:
          isDragging && buttonAllowDropFiles
            ? `${_theme.colorBrandBackground2}25`
            : "transparent",
        backdropFilter:
          isDragging && buttonAllowDropFiles ? "blur(10px)" : "none",
      }}
    >
      {buttonVisible ? (
        <FluentProvider theme={_theme}>
          {buttonShowSecondaryContent ? (
            <CompoundButton
              onClick={handleButtonClick}
              {..._buttonProps}
              secondaryContent={buttonSecondaryContent}
            >
              {buttonText}
            </CompoundButton>
          ) : (
            <Button onClick={handleButtonClick} {..._buttonProps}>
              {buttonText}
            </Button>
          )}
        </FluentProvider>
      ) : null}
      <input
        ref={importFileRef}
        multiple={buttonAllowMultipleFiles}
        type="file"
        accept={buttonAllowedFileTypes}
        onChange={(e) => {
          void onFileChange(e);
        }}
        style={{ display: "none" }}
      />
      {isDragging && buttonAllowDropFiles && (
        <Caption1>{buttonAllowDropFilesText}</Caption1>
      )}
    </div>
  );
};
