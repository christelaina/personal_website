import React, { useState, useEffect } from "react";
import "./picturebox.css";

const Picturebox = () => {
  const imagesArray = [
    "mushroom_01.jpg",
    "moment_08.jpg",
    "moment_07.jpg",
    "moment_06.jpg",
    "moment_05.jpg",
    "moment_04.jpg",
    "moment_03.jpg",
    "moment_02.jpg",
    "moment_01.jpg",
    "kayak_01.jpg",
    "flower_02.jpg",
    "flower_01.jpg",
    "cat_03.jpg",
    "cat_02.jpg",
    "cat_01.jpg",
    "bird_02.jpg",
    "bird_01.jpg",
  ];

  const [imageIndex, setImageIndex] = useState(0);

  useEffect(() => {
    const initialImage = Math.floor(Math.random() * imagesArray.length);
    setImageIndex(initialImage);
  }, [imagesArray.length]);

  const changeImage = () => {
    const randomIndex = Math.floor(Math.random() * imagesArray.length);
    setImageIndex(randomIndex);
  };

  return (
    <>
      <div className="picturebox-container">
        <div className="picturebox">
          <img
            src={require(`../assets/images/${imagesArray[imageIndex]}`)}
            onClick={changeImage}
            alt=""
            className="picturebox-img"
          />
        </div>
      </div>
    </>
  );
};

export default Picturebox;
