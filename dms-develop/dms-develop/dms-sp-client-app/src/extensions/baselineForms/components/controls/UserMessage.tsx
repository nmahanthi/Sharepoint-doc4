/* eslint-disable @typescript-eslint/explicit-function-return-type */
/**

AVAILABLE PROPS:
+------------+--------------------------+---------------+---------------------------------------------------------+
| Name       | Type                     | Default Value | Description                                             |
+------------+--------------------------+---------------+---------------------------------------------------------+
| type       | (error | success | info) |               | Determines the style of the message display.            |
| (required) |                          |               | (Error = red box, Success = green box, etc)             |
+------------+--------------------------+---------------+---------------------------------------------------------+
| message    | string                   |               | Supply a message to display. You can type this directly |
| (required) |                          |               | or reference props/state from the parent of course.     |
+------------+--------------------------+---------------+---------------------------------------------------------+
| visible    | boolean                  | true          | Whether the component should be displayed or not.       |
|            |                          |               | Typically this will be handled by a state or prop       |
|            |                          |               | from the parent component. You can also handle the      |
|            |                          |               | logic of showing the component in the parent directly.  |
+------------+--------------------------+---------------+---------------------------------------------------------+
| className  | string                   | ''            | Supply a custom className.                              |
|            |                          |               | (Cannot be 'notice', 'error', 'success' or 'info'.)     |
+------------+--------------------------+---------------+---------------------------------------------------------+

*/

import * as React from 'react';
import { Icon } from '@fluentui/react';
import styles from './Controls.module.scss';

export interface IUserMessageProps {
  type: ('error' | 'success' | 'info');
  message: string;
  visible?: boolean;
  className?: Exclude<string, 'notice' | 'error' | 'success' | 'info'>;
}

export interface IUserMessageState {
  showMessage?: boolean;
  iconName: string;
}

export default class UserMessage extends React.Component<IUserMessageProps, IUserMessageState> {
  constructor(props: IUserMessageProps | Readonly<IUserMessageProps>) {
    super(props);
    this.state = {
      showMessage: true,
      iconName: ''
    };
  }

  public componentDidMount = async () => {
    const iconName = await this.getIconName(this.props.type);
    this.setState({
      showMessage: 'visible' in this.props ? this.props.visible : true,
      iconName
    });
    this.scrollToMessage();
  }

  public componentDidUpdate = async (newProps: Readonly<IUserMessageProps> & Readonly<{ children?: React.ReactNode; }>) => {
    if (newProps !== this.props) {
      const iconName = await this.getIconName(this.props.type);
      this.setState({
        showMessage: 'visible' in this.props ? this.props.visible : true,
        iconName
      });
    }
  }

  private scrollToMessage = () => {
    const header = document.getElementById('panelHeader');
    if (header !== null) {
      header.scrollIntoView({ behavior: 'smooth', block: 'start' });
      return;
    }
    const msg = document.getElementById('userMessage');
    if (msg !== null) msg.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  public getIconName = (type: string) => {
    // Can add additional icons with additional 'if' statements.
    let iconName: string = '';

    if (type === 'error') {
      iconName = 'Error';
    }
    if (type === 'success') {
      iconName = 'CheckMark';
    }
    if (type === 'info') {
      iconName = 'Info';
    }

    return iconName;
  }

  public render(): React.ReactElement<IUserMessageProps> {
    return (
      <React.Fragment>
        {this.state.showMessage &&
          // eslint-disable-next-line no-useless-concat
          <div className={`${styles.notice} ${this.props.type === 'info' ? styles.info : this.props.type === 'error' ? styles.error : styles.success}`} id="userMessage">
            {this.state.iconName !== '' ? <Icon iconName={this.state.iconName} /> : <React.Fragment />}
            <span>&nbsp;{this.props.message}</span>
          </div>
        }
      </React.Fragment>
    );
  }
}