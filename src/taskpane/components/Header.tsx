import * as React from "react";

export interface HeaderProps {
  title: string;
  logo: string;
  message: string;
}

export default class Header extends React.Component<HeaderProps> {
  render() {
    const { title, logo, message } = this.props;

    return (
      <section className="ms-welcome__header ms-bgColor-neutralLighter ms-u-fadeIn500">
        <img width="auto" height="50" src={logo} alt={title} title={title} />
        <h1 style={{ fontSize: "1.5rem" }} className="ms-fontSize-su ms-fontWeight-light ms-fontColor-neutralPrimary">
          {message}
        </h1>
      </section>
    );
  }
}
