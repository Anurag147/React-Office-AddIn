import * as React from "react";
import { Button } from "@fluentui/react-components";

interface HighlightProps {
  highlight: (text: string) => void;
}

const Highlight = (props: HighlightProps) => {
  const highlight = async () => {
    await props.highlight("My name is Anurag");
  };

  return (
    <>
      <Button appearance="primary" disabled={false} size="large" onClick={highlight}>
        Highlight text
      </Button>
    </>
  );
};

export default Highlight;
