/* eslint-disable no-undef */
import React, { useEffect, useState } from "react";
import { DefaultButton } from "@fluentui/react";

/* global  Office  */

interface TekstiAlyButtonProps {
  name: string;
  instructions: string;
}

function TekstiAlyButton({ name, instructions }: TekstiAlyButtonProps) {
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

    const data = {
      messages: [
        {
          role: "system",
          content: "Olet kohtelias ja avulias tekoälyassistentti",
        },
        {
          role: "user",
          content: `${instructions}: ${text}`,
        },
      ],
      max_tokens: 4000,
      temperature: 0,
    };
    // eslint-disable-next-line no-undef
    fetch(
      `https://gptlab.openai.azure.com/openai/deployments/chatgpt-deployment/chat/completions?api-version=2023-03-15-preview`,
      {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "api-key": "1dfa1eaeb7f44b708d016307d7c59f39",
        },
        body: JSON.stringify(data),
      }
    )
      .then((result) => result.json())
      .then((data) => {
        Office.context.mailbox.item.body.setAsync(data.choices[0].message.content, {
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

export default TekstiAlyButton;
