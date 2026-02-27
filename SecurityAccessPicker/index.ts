import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as React from "react";
import { BentoGrid } from "./BentoGrid";
import "./BentoGrid.css";

import { FluentProvider } from "@fluentui/react-provider";
import { webLightTheme } from "@fluentui/react-theme";
import { RendererProvider, createDOMRenderer } from "@griffel/react";
import { SSRProvider } from "@fluentui/react-utilities";

interface FluentRendererProviderProps {
  renderer: ReturnType<typeof createDOMRenderer>;
  children?: React.ReactNode;
}

const FluentRendererProvider = RendererProvider as React.ComponentType<FluentRendererProviderProps>;

export class SecurityAccessPicker
  implements ComponentFramework.ReactControl<IInputs, IOutputs>
{
  private static readonly renderer = createDOMRenderer();
  private notifyOutputChanged: (() => void) | undefined;
  private rootId = `security-access-picker-root-${Math.random().toString(36).slice(2)}`;

  public init(
    context: ComponentFramework.Context<IInputs>,
    notifyOutputChanged: () => void,
    state: ComponentFramework.Dictionary
  ): void {
    this.notifyOutputChanged = notifyOutputChanged;
    context.mode.trackContainerResize(true);
  }

  public updateView(context: ComponentFramework.Context<IInputs>): React.ReactElement {
    const allocatedHeight = context.mode.allocatedHeight;
    const viewportHeight = window.visualViewport?.height && window.visualViewport.height > 0
      ? window.visualViewport.height
      : window.innerHeight > 0
      ? window.innerHeight
      : 0;
    const rootElement = document.getElementById(this.rootId);
    const topOffset = rootElement ? Math.max(0, rootElement.getBoundingClientRect().top) : 0;
    const viewportAvailableHeight = viewportHeight > 0 ? Math.max(0, viewportHeight - topOffset) : 0;
    const clampedHeightPx =
      allocatedHeight > 0 && viewportAvailableHeight > 0
        ? Math.min(allocatedHeight, viewportAvailableHeight)
        : allocatedHeight > 0
        ? allocatedHeight
        : viewportAvailableHeight;
    const height = clampedHeightPx > 0 ? `${clampedHeightPx}px` : "100dvh";
    const containerStyle: React.CSSProperties = {
      width: "100%",
      height,
      maxHeight: height,
      minWidth: 0,
      minHeight: 0,
      overflow: "hidden",
      display: "flex",
      flexDirection: "column",
    };

    return React.createElement(
      FluentRendererProvider,
      { renderer: SecurityAccessPicker.renderer },
      React.createElement(SSRProvider, null, React.createElement(
        FluentProvider,
        { id: this.rootId, theme: webLightTheme, style: containerStyle },
        React.createElement(BentoGrid, { context })
      ))
    );
  }

  public getOutputs(): IOutputs {
    return {};
  }

  public destroy(): void {
    this.notifyOutputChanged = undefined;
  }
}
