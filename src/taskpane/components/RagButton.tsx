/* eslint-disable no-undef */
import React, { useEffect, useState } from "react";
import { DefaultButton } from "@fluentui/react";

/* global  Office  */

interface RagButtonProps {
  name: string;
  instructions: string;
}

function RagButton({ name, instructions }: RagButtonProps) {
  const [text, setText] = useState("");
  const [searchCount, setSearchCount] = useState(0);

  const click = async () =>
    Office.context.mailbox.item.body.getAsync(
      "text",
      { asyncContext: "This is passed to the callback" },
      function callback(result) {
        // Do something with the result.
        setText(result.value);
        setSearchCount((prev) => prev + 1);
      }
    );

  useEffect(() => {
    if (searchCount === 0) return;

    const data = { query: `${instructions}: ${text}`, n_documents: 6 };

    // eslint-disable-next-line no-undef
    fetch(`https://rag-wrapper.azurewebsites.net/api/HttpTrigger1`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-functions-key": "FDf8ZMrJ6-LnbygIsvOI9mwfUUP_tvqm__084xg7eoevAzFuThEmMQ==",
      },
      body: JSON.stringify(data),
    })
      .then((result) => result.json())
      .then((data) => {
        Office.context.mailbox.item.body.setAsync(data.answer, {
          coercionType: Office.CoercionType.Text,
        });
      });
  }, [searchCount]);

  return (
    <>
      <DefaultButton
        style={{
          backgroundColor: "#E00034",
          color: "white",
          fontSize: "rem",
          borderRadius: "2px",
          boxShadow: "rgba(50, 50, 93, 0.25) 0px 2px 5px -1px, rgba(0, 0, 0, 0.3) 0px 1px 3px -1px",
          border: "none",
        }}
        className="ms-welcome__action"
        iconProps={{ iconName: "ChevronRight" }}
        onClick={click}
      >
        {name}
      </DefaultButton>
    </>
  );
}

export default RagButton;
