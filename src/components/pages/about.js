import InteractiveSVG from "../interactivesvg";
import Picturebox from "../picturebox";
import Sidebar from "../sidebar";
import "../../App.css"

const About = () => {
  return (
    <div className="parent-content">
      <Sidebar />
      <div className="content">
        <div className="wrap">
          <div>
            <h1>about me</h1>
            <p>I'm a computer science student. I love to learn about tech, mycology, and the nature of things. On my spare time, I'm collecting new hobbies and pass times.</p>
            <InteractiveSVG/>
          </div>
          <Picturebox/>
        </div>
        
      </div>
    </div>
  );
};

export default About;
