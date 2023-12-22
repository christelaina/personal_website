import Picturebox from "../picturebox";
import Sidebar from "../sidebar";
import "../../App.css";
import Footbar from "../footbar";

const About = () => {
  return (
    <div className="parent-content">
      <Sidebar />
      <div className="content">
        <div className="wrap">
          <div>
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
        <Footbar/>
      </div>
    </div>
  );
};

export default About;
