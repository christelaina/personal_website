import React, { useState, useEffect } from "react";
import "./picturebox.css";

const Picturebox = () => {
  const imagesArray = [
    {
      theme: "birds",
      images: ["IMG_7871.jpg", "IMG_8165.jpg", "IMG_9379.jpg"],
      text: "sometimes I use the Merlin Bird ID app on my phone to 'capture' singing birds. birds have a sweet fragility, always fleeting. I'm grateful to identify one when they allow me near."
    },
    {
      theme: "cats",
      images: [
        "IMG_1286.jpg",
        "IMG_3816.jpg",
        "IMG_5165.jpg",
        "IMG_7679.jpg",
        "IMG_8545.jpg",
      ],
      text: "the cream coloured cat is named Natto. his coat colour resembles the Japanese dish of fermented soy beans. the dark coated cat is Maru, short for Marimo, meaning mossball. she's always sneezing booger bombs."
    },
    { theme: "flowers", images: ["IMG_7676.jpg", "IMG_8097.jpg"], text: "always press or hang dry flowers for sentimental preservation." },
    { theme: "kayaking", images: ["IMG_6507.jpg", "IMG_6646.jpg"], text: "first time back country camping. would highly recommend." },
    {
      theme: "moments",
      images: [
        "IMG_7452.jpg",
        "IMG_3885.jpg",
        "IMG_4472.jpg",
        "IMG_5321.jpg",
        "IMG_5411.jpg",
        "IMG_6203.jpg",
        "IMG_6262.jpg",
        "IMG_7158.jpg",
        "IMG_7453.jpg",
        "IMG_7689.jpg",
        "IMG_8075.jpg",
        "IMG_8782.jpg",
      ],
      text: "pictures with no significant content, but are moments with deeper meaning. capturing the way I engage with the everyday world, and creating portals."
    },
    {
      theme: "mushrooms",
      images: [
        "IMG_4010.jpg",
        "IMG_1435.jpg",
        "PXL_20220716_193406661.jpg",
        "PXL_20220716_210034470.jpg",
      ],
      text: "mushrooms found on hikes. a good sign of a healthy forest."
    },
  ];
  const [randomThemeIndex, setRandomThemeIndex] = useState(0);
  const [randomImageIndex, setRandomImageIndex] = useState(0);

  useEffect(() => {
    const initialTheme = Math.floor(Math.random() * imagesArray.length);
    setRandomThemeIndex(initialTheme);
  }, [imagesArray.length]);

  const changeImage = () => {
    const selectedTheme = imagesArray[randomThemeIndex];
    const selectedImage = Math.floor(
      Math.random() * selectedTheme.images.length
    );
    setRandomImageIndex(selectedImage);
  };

  return (
    <>
      
      <div className="picturebox">
        <img
          src={require(`../assets/${imagesArray[randomThemeIndex].theme}/${imagesArray[randomThemeIndex].images[randomImageIndex]}`)}
          onClick={changeImage}
          alt=""
        />
      </div>
      <div className="textbox">
        <h3>the art of noticing</h3>
        <p>{imagesArray[randomThemeIndex].text}</p>
      </div>
    </>
  );
};

export default Picturebox;
