import * as React from 'react';
import styles from './HelloWorld.module.scss';
import { IHelloWorldProps } from './IHelloWorldProps';
import { escape } from '@microsoft/sp-lodash-subset';

import { FileTypeIcon, ApplicationType, IconType, ImageSize } from "@pnp/spfx-controls-react/lib/FileTypeIcon";

export default class HelloWorld extends React.Component<IHelloWorldProps, {}> {
  public render(): React.ReactElement<IHelloWorldProps> {
    return (
      <div>
        <FileTypeIcon type={IconType.font} application={ApplicationType.Excel} />
        <FileTypeIcon type={IconType.font} path="https://contoso.sharepoint.com/documents/filename.docx" />
        <FileTypeIcon type={IconType.font} application={ApplicationType.ASPX} />
        <FileTypeIcon type={IconType.font} application={ApplicationType.Code} />
        <br />
        <FileTypeIcon type={IconType.image} application={ApplicationType.Word} />
        <FileTypeIcon type={IconType.image} path="https://contoso.sharepoint.com/documents/filename.xlsx" />
        <br />
        <FileTypeIcon type={IconType.image} size={ImageSize.small} application={ApplicationType.Excel} />
        <FileTypeIcon type={IconType.image} size={ImageSize.medium} application={ApplicationType.Excel} />
        <FileTypeIcon type={IconType.image} size={ImageSize.large} application={ApplicationType.Excel} />
      </div>
    );
  }
}
