import "../../App.css";
import Footbar from "../footbar";
import Navbar from "../navbar";

const Home = () => {
  return (
    <>
      <div className="center">
        <div className="parent-content">
          <Navbar />
          <div className="content">
            <div className="blob"></div>
          </div>
          <Footbar />
        </div>
      </div>
    </>
  );
};

export default Home;
