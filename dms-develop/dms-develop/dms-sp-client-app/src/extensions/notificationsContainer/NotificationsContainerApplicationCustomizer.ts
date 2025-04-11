import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';

import ReactDOM from 'react-dom';
import React from 'react';
import NotificationsContainer, { INotificationsContainerProps } from './components/NotificationsContainer';
import { NotificationService } from '../../service/NotificationService';

const LOG_SOURCE: string = 'NotificationsContainerApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface INotificationsContainerApplicationCustomizerProperties {
  // This is an example; replace with your own property
  testMessage: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class NotificationsContainerApplicationCustomizer
  extends BaseApplicationCustomizer<INotificationsContainerApplicationCustomizerProperties> {
  private _topPlaceholder: PlaceholderContent | undefined;
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized`);
    this.context.placeholderProvider.changedEvent.add(this, this._renderPlaceHolders);
    return Promise.resolve();
  }
  private _renderPlaceHolders(): void {
    if (!this._topPlaceholder) {
      this._topPlaceholder = this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Top,
        { onDispose: this._onDispose.bind(this) }
      );

      if (!this._topPlaceholder) {
        return;
      }

      this.context.serviceScope.whenFinished(() => {
        if (this._topPlaceholder?.domElement) {
          const notificationsContainer: React.ReactElement<{}> =
            React.createElement(NotificationsContainer, {
              notificationService: this.context.serviceScope.consume(NotificationService.serviceKey)
            } as INotificationsContainerProps);
          ReactDOM.render(notificationsContainer, this._topPlaceholder.domElement);
        }
      });
    }
  }
  private _onDispose(placeholderContent: PlaceholderContent): void {
    if (placeholderContent.domElement) {
      ReactDOM.unmountComponentAtNode(placeholderContent.domElement);
    }
  }
}
