import * as React from "react";
import PropTypes from "prop-types";
import { DefaultButton } from "@fluentui/react";
//import Header from "./Header";
import HeroList from "./HeroList";
import Progress from "./Progress";

/* global Word, require */

export default class App extends React.Component {
  constructor(props, context) {
    super(props, context);
    this.state = {
      listItems: [],
    };
  }

  componentDidMount() {
    this.setState({
      listItems: [
        // {
        //   icon: "Ribbon",
        //   primaryText: "Achieve more with Office integration",
        // },
        // {
        //   icon: "Unlock",
        //   primaryText: "Unlock features and functionality",
        // },
        // {
        //   icon: "Design",
        //   primaryText: "Create and visualize like a pro",
        // },
      ],
    });
  }

  click = async () => {
    return Word.run(async (context) => {
      /**
       * Insert your Word code here
       *
       */
      const url = "https://jsonplaceholder.typicode.com/todos/1";
      const response = await fetch(url);

      //Expect that status code is in 200-299 range
      if (!response.ok) {
        throw new Error(response.statusText);
      }
      const data = await response.json();
      console.log("Response: " + data.id);

      let doc = context.document;
      let paragraphs = context.document.body.paragraphs;
      paragraphs.load("$none");

      await context.sync();

      // let contentControl = paragraphs.items[0].insertContentControl();
      // let contentControl2 = paragraphs.items[1].insertContentControl();
      // let contentControl3 = paragraphs.items[2].insertContentControl();
      // contentControl.tag = "amount";
      // contentControl2.tag = "amount";
      // contentControl3.tag = "amount";

      let contentControls = context.document.contentControls.getByTag("amount");
      contentControls.load("text");

      await context.sync();

      //copied from docs, will refactor
      let a = Math.random();
      for (let i = 0; i < contentControls.items.length; i++) {
        contentControls.items[i].insertText((data.id * a * 1000).toString(), "Replace");
      }

      await context.sync();

      // doc.body.load(["style"]);
      // await context.sync();
      // doc.body.style = "Title";

      // // insert a paragraph at the end of the document.
      // const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);

      // // change the paragraph color to blue.
      // paragraph.font.color = "blue";

      await context.sync();
    });
  };

  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress
          title={title}
          logo={require("./../../../assets/logo-filled.png")}
          message="Please sideload your addin to see app body."
        />
      );
    }

    return (
      <div className="ms-welcome">
        {/* <Header logo={require("./../../../assets/logo-filled.png")} title={this.props.title} message="Welcome" /> */}
        <HeroList message="Update the deal info for this document" items={this.state.listItems}>
          <p className="ms-font-l"></p>
          <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={this.click}>
            Update deal info
          </DefaultButton>
        </HeroList>
      </div>
    );
  }
}

App.propTypes = {
  title: PropTypes.string,
  isOfficeInitialized: PropTypes.bool,
};
