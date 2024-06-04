import * as React from 'react';
import styles from './PropertyPaneIcons.module.scss';
import type { IPropertyPaneIconsProps } from './IPropertyPaneIconsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Icon } from '@fluentui/react';

export interface IState{
  icon:any;
}

export default class PropertyPaneIcons extends React.Component<IPropertyPaneIconsProps, IState, {}> {
  constructor(props: IPropertyPaneIconsProps, state: IState) {    
    super(props);    
    this.state = {icon: ""};    
  } 
  public render(): React.ReactElement<IPropertyPaneIconsProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className={`${styles.propertyPaneIcons} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
          <h2>Well done, {escape(userDisplayName)}!</h2>
          <div>{environmentMessage}</div>
          <div>Web part property value: <strong>{escape(description)}</strong></div>
        </div>
        <Icon style={{fontSize:'24px'}} iconName={this.props.iconPicker1} className="ms-IconExample" />
        <Icon style={{fontSize:'24px'}} iconName={this.props.iconPicker2} className="ms-IconExample" />
        <Icon style={{fontSize:'24px'}} iconName={this.props.iconPicker3} className="ms-IconExample" />

        <div>
          <h3>Welcome to SharePoint Framework!</h3>
          <p>
            The SharePoint Framework (SPFx) is a extensibility model for Microsoft Viva, Microsoft Teams and SharePoint. It&#39;s the easiest way to extend Microsoft 365 with automatic Single Sign On, automatic hosting and industry standard tooling.
          </p>
          <h4>Learn more about SPFx development:</h4>
          <ul className={styles.links}>
            <li><a href="https://aka.ms/spfx" target="_blank" rel="noreferrer">SharePoint Framework Overview</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-graph" target="_blank" rel="noreferrer">Use Microsoft Graph in your solution</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-teams" target="_blank" rel="noreferrer">Build for Microsoft Teams using SharePoint Framework</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-viva" target="_blank" rel="noreferrer">Build for Microsoft Viva Connections using SharePoint Framework</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-store" target="_blank" rel="noreferrer">Publish SharePoint Framework applications to the marketplace</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-api" target="_blank" rel="noreferrer">SharePoint Framework API reference</a></li>
            <li><a href="https://aka.ms/m365pnp" target="_blank" rel="noreferrer">Microsoft 365 Developer Community</a></li>
          </ul>
        </div>
      </section>
    );
  }
}
