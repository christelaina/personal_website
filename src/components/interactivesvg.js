import React, { useRef, useState } from "react";
import "./interactivesvg.css";

const InteractiveSVG = () => {
  const svgRef = useRef(null);
  const [isDragging, setDragging] = useState(false);
  const [dragStart, setDragStart] = useState({ x: 100, y: 250 });
  const [position, setPosition] = useState({ x: 100, y: 250 });

  const handleMouseDown = (e) => {
    setDragging(true);
    setDragStart({ x: e.clientX, y: e.clientY });
  };

  const handleMouseMove = (e) => {
    if (!isDragging) return;

    const deltaX = e.clientX - dragStart.x;
    const deltaY = e.clientY - dragStart.y;

    setPosition({
      x: position.x + deltaX,
      y: position.y + deltaY,
    });

    setDragStart({ x: e.clientX, y: e.clientY });
  };

  const handleMouseUp = () => {
    setDragging(false);
  };

  return (
    <svg
      ref={svgRef}
      width="350"
      height="350"
      className="free-roam"
      onMouseDown={handleMouseDown}
      onMouseMove={handleMouseMove}
      onMouseUp={handleMouseUp}
    >
      <circle
        cx={position.x}
        cy={position.y}
        r="100"
        fill="#283106"
        className="dot"
      />
    </svg>
  );
};

export default InteractiveSVG;
