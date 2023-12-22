import React, { useState, useEffect } from "react";
import {
  FirstFrost,
  Spring,
  Ivy,
  Pebble,
  Linen,
  Serpentine,
} from "./../assets/testsvg";
import "./animate.css";

const svgList = [FirstFrost, Spring, Ivy, Pebble, Linen, Serpentine];

const AnimatedSVG = () => {
  const [currentSVGIndex, setcurrentSVGIndex] = useState(0);

  useEffect(() => {
    const intervalId = setInterval(() => {
        setcurrentSVGIndex((prevIndex) => (prevIndex + 1) % svgList.length);
    }, 5000);

    return () => clearInterval(intervalId);
  }, []);

  const CurrentSVG = svgList[currentSVGIndex];
  return (
    <div className="">
      <CurrentSVG />
    </div>
  );
};

export default AnimatedSVG;
