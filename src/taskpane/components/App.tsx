import * as React from "react";
import Header from "./Header";
import Progress from "./Progress";
import TekstiAlyButton from "./TekstiAlyButton";

/* global require */

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export default class App extends React.Component<AppProps> {
  constructor(props, context) {
    super(props, context);
    this.state = {
      listItems: [],
    };
  }

  componentDidMount() {
    this.setState({
      listItems: [
        {
          icon: "Ribbon",
          primaryText: "Achieve more with Office integration",
        },
        {
          icon: "Unlock",
          primaryText: "Unlock features and functionality",
        },
        {
          icon: "Design",
          primaryText: "Create and visualize like a pro",
        },
      ],
    });
  }

  click = async () => {
    /**
     * Insert your Outlook code here
     */
  };

  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress
          title={title}
          logo={require("./../../../assets/efima-logo.png")}
          message="Please sideload your addin to see app body."
        />
      );
    }

    return (
      <div className="ms-welcome">
        <Header
          logo={require("./../../../assets/efima-logo.png")}
          title={this.props.title}
          message="tekstiÄly (Preview 44/23)"
        />
        <div
          style={{
            padding: "30px 10px 30px 50px",
            display: "flex",
            flexDirection: "column",
            gap: "20px",
            alignItems: "flex-start",
          }}
        >
          <p style={{ fontSize: "2rem" }} className="ms-fontSize-su ms-fontWeight-light ms-fontColor-neutralPrimary">
            Pikatoiminnot
          </p>
          <div style={{ display: "flex", flexDirection: "column", gap: "20px" }}>
            <TekstiAlyButton name="Käännä englanniksi" instructions="Käännä seuraava teksti englanniksi" />
            <TekstiAlyButton name="Käännä suomeksi" instructions="Käännä seuraava teksti suomeksi" />
            <TekstiAlyButton
              name="Tarkista oikeinkirjoitus"
              instructions="Tarkista seuraavan tekstin oikeinkirjoitus ja anna vastaukseksi parannettu versio tekstistä"
            />
            <TekstiAlyButton
              name="Muuta ystävällisemmäksi"
              instructions="Tee seuraavasta tekstistä ystävällisempi versio"
            />
            <TekstiAlyButton name="Muuta asiallisemmaksi" instructions="Tee seuraavasta tekstistä asiallisempi" />
          </div>
        </div>
      </div>
    );
  }
}
