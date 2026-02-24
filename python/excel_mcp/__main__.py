from .mcp_server import _kill_previous_instance, mcp


if __name__ == "__main__":
  _kill_previous_instance()
  mcp.run()
