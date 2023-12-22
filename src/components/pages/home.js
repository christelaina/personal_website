import "../../App.css";
import Nameplate from "../nameplate";
import Sidebar from "../sidebar";

const Home = () => {
  return (
    <>
      <div className="parent-content">
        <Sidebar />
        <div className="content">
          <Nameplate />
        </div>
      </div>
    </>
  );
};

export default Home;
