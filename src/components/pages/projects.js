import Navbar from "../navbar";
import "../../App.css";
import Footbar from "../footbar";

const Projects = () => {
  return (
    <>
      <div className="center">
        <div className="parent-content">
          <Navbar />
          <div className="content">
            <h1>projects</h1>
          </div>
          <Footbar />
        </div>
      </div>
    </>
  );
};

export default Projects;
