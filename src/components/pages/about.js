import Picturebox from "../picturebox";
import Navbar from "../navbar";
import "../../App.css";
import Footbar from "../footbar";

const About = () => {
  return (
    <>
      <div className="center">
        <div className="parent-content">
          <Navbar />
          <div className="wrap">
            <div className="group">
              <h1>about me</h1>
              <p>
                I'm a computer science student. I love to learn about tech,
                mycology, and the nature of things. On my spare time, I'm
                collecting new hobbies and pass times.
              </p>
              <svg width="350" height="350">
                <circle cx="100" cy="250" r="200" fill="#283106" />
              </svg>
            </div>
            <Picturebox />
          </div>
          <Footbar />
        </div>
      </div>
    </>
  );
};

export default About;
