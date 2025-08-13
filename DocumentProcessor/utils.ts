import { Button, type ButtonProps } from "@fluentui/react-components";

export enum ButtonLoadingStateEnum {
  Initial = "initial",
  Loading = "loading",
  Loaded = "loaded"
}

//
enum ButtonAppearanceEnum {
    primary = 0,
    secondary = 1,
    outline = 2,
    subtle = 3,
    transparent = 4
}

//
enum ButtonShapeEnum {
    rounded = 0,
    circular = 1,
    square = 2
}

//
enum ButtonIconPositionEnum {
    before = 0,
    after = 1
}

//
enum ButtonSizeEnum {
    small = 0,
    medium = 1,
    large = 2
}

//
enum ButtonAlign {
    left = 0,
    center = 1,
    right = 2,
    justify = 3
}
const getButtonAlign = (align: string): string => {
    const parsedEnum = Number.parseInt(align);
    if (isNaN(parsedEnum) || !(parsedEnum in ButtonAlign)) {
        throw new Error(`Invalid button align: ${align}`);
    }

    switch (ButtonAlign[parsedEnum]) {
        case ButtonAlign[ButtonAlign.left]:
            return "flex-start";
        case ButtonAlign[ButtonAlign.center]:
            return "center";
        case ButtonAlign[ButtonAlign.right]:
            return "flex-end";
        case ButtonAlign[ButtonAlign.justify]:
            return "stretch";
        default:
            return "center";
    }
}

//
enum ButtonFontWeight {
    bold = 0,
    lighter = 1,
    normal = 2,
    semibold = 3
}
const getButtonFontWeight = (weight: string): string => {
    const parsedEnum = Number.parseInt(weight);
    if (isNaN(parsedEnum) || !(parsedEnum in ButtonFontWeight)) {
        throw new Error(`Invalid button font weight: ${weight}`);
    }

    switch (ButtonFontWeight[parsedEnum]) {
        case ButtonFontWeight[ButtonFontWeight.bold]:
            return "bold";
        case ButtonFontWeight[ButtonFontWeight.semibold]:
            return "semibold";
        case ButtonFontWeight[ButtonFontWeight.normal]:
            return "normal";
        case ButtonFontWeight[ButtonFontWeight.lighter]:
            return "lighter";
        default:
            return "normal";
    }
}

//
export const getButtonAppearance = (appearance: string): ButtonProps["appearance"] => {
    const parsedEnum = Number.parseInt(appearance);
    if (isNaN(parsedEnum) || !(parsedEnum in ButtonAppearanceEnum)) {
        throw new Error(`Invalid button appearance: ${appearance}`);
    }
    return ButtonAppearanceEnum[parsedEnum] as ButtonProps["appearance"];
}

//
export const getButtonIconPosition = (iconPosition: string): ButtonProps["iconPosition"] => {
    const parsedEnum = Number.parseInt(iconPosition);
    if (isNaN(parsedEnum) || !(parsedEnum in ButtonIconPositionEnum)) {
        throw new Error(`Invalid button icon position: ${iconPosition}`);
    }
    return ButtonIconPositionEnum[parsedEnum] as ButtonProps["iconPosition"];
}

//
export const getButtonShape = (shape: string): ButtonProps["shape"] => {
    const parsedEnum = Number.parseInt(shape);
    if (isNaN(parsedEnum) || !(parsedEnum in ButtonShapeEnum)) {
        throw new Error(`Invalid button shape: ${shape}`);
    }
    return ButtonShapeEnum[parsedEnum] as ButtonProps["shape"];
}

//
export const getButtonSize = (size: string): ButtonProps["size"] => {
    const parsedEnum = Number.parseInt(size);
    if (isNaN(parsedEnum) || !(parsedEnum in ButtonSizeEnum)) {
        throw new Error(`Invalid button size: ${size}`);
    }
    return ButtonSizeEnum[parsedEnum] as ButtonProps["size"];
}

//
export const getButtonStyle = (width: number, height: number, align: string, weight: string): ButtonProps["style"] => {
    return { 
        width: `${width}px`, 
        height: `${height}px`,
        justifyContent: `${getButtonAlign(align)}`,
        fontWeight: `${getButtonFontWeight(weight)}`,
    };
}

//
export const getDisplayMode = (
    loadingState: string,
    buttonIsDisabled: boolean
): boolean => {
    if (loadingState === "loading") {
        return true; // always disabled while loading
    }
    return buttonIsDisabled;
};

