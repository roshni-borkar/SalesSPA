/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable no-prototype-builtins */
import * as React from "react";
import * as AdaptiveCards from "adaptivecards";

interface AdaptiveCardRendererProps {
  card: any;
  data?: any;
  onSubmit?: (data: any) => void;
}

const AdaptiveCardRenderer: React.FC<AdaptiveCardRendererProps> = ({ card, data, onSubmit }) => {
  const cardRef = React.useRef<HTMLDivElement>(null);

  React.useEffect(() => {
    const adaptiveCard = new AdaptiveCards.AdaptiveCard();
    adaptiveCard.parse(card);

    adaptiveCard.onExecuteAction = (action) => {
      const inputs = adaptiveCard.getAllInputs().reduce((acc: any, input) => {
        if (input.id !== undefined) {
          acc[input.id] = input.value;
        }
        return acc;
      }, {});
      onSubmit?.(inputs);
    };

    if (data) {
      adaptiveCard.getAllInputs().forEach((input: any) => {
        if (input.id && data.hasOwnProperty(input.id)) {
          input.value = data[input.id];
        }
      });
    }
    const renderedCard = adaptiveCard.render();

    if (cardRef.current && renderedCard) {
      cardRef.current.innerHTML = "";
      cardRef.current.appendChild(renderedCard);
    }
  }, [card, data]);

  return <div ref={cardRef} />;
};

export default AdaptiveCardRenderer;
