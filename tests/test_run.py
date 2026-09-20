"""Launcher checks that do not start servers or require installed dependencies."""

from __future__ import annotations

import contextlib
import errno
import inspect
import io
import os
import sys
import unittest
from pathlib import Path
from unittest.mock import MagicMock, call, patch

import run


class RunTests(unittest.TestCase):
    def launch(self, arguments=(), environment=None, selected_ports=(8000, 5173)):
        args = run.parse_args(list(arguments))
        backend = MagicMock()
        backend.poll.return_value = 0
        frontend = MagicMock()
        port_signature = inspect.signature(run.available_port)
        with (
            patch.dict(os.environ, environment or {}, clear=True),
            patch.object(run, "parse_args", return_value=args),
            patch.object(run, "select_runtime", return_value=Path(sys.executable)),
            patch.object(run, "load_env_file"),
            patch.object(run, "available_port", side_effect=selected_ports) as probe,
            patch.object(run.shutil, "which", return_value="npm"),
            patch.object(Path, "is_dir", return_value=True),
            patch.object(run.subprocess, "Popen", side_effect=[backend, frontend]) as popen,
            patch.object(run, "stop"),
            contextlib.redirect_stdout(io.StringIO()),
        ):
            self.assertEqual(run.main(), 0)
        probes = []
        for invocation in probe.call_args_list:
            bound = port_signature.bind(*invocation.args, **invocation.kwargs)
            bound.apply_defaults()
            probes.append(bound.arguments)
        return popen.call_args_list, probes

    @staticmethod
    def option(command, flag, default=None):
        return command[command.index(flag) + 1] if flag in command else default

    def test_serve_and_lan_arguments_accept_independent_ports(self):
        args = run.parse_args([
            "serve", "--lan", "--frontend-port", "5200", "--backend-port", "8100",
        ])
        self.assertTrue(args.lan)
        self.assertEqual(args.frontend_port, 5200)
        self.assertEqual(args.backend_port, 8100)
        self.assertFalse(run.parse_args([]).lan)
        self.assertEqual(run.parse_args(["--backend-port", "8100"]).backend_port, 8100)

    def test_lan_overrides_env_host_and_uses_actual_selected_ports(self):
        processes, probes = self.launch(
            ["serve", "--lan", "--backend-port", "8100", "--frontend-port", "5200", "--no-reload"],
            {"BACKEND_HOST": "127.0.0.1", "BACKEND_PORT": "9000", "FRONTEND_PORT": "9001"},
            selected_ports=(8101, 5201),
        )
        backend_command, frontend_command = [item.args[0] for item in processes]
        self.assertEqual(self.option(backend_command, "--host"), "0.0.0.0")
        self.assertEqual(self.option(frontend_command, "--host"), "0.0.0.0")
        self.assertEqual(self.option(backend_command, "--port"), "8101")
        self.assertEqual(self.option(frontend_command, "--port"), "5201")
        self.assertIn("--strictPort", frontend_command)
        self.assertNotIn("--reload", backend_command)
        self.assertEqual(processes[1].kwargs["env"]["VITE_BACKEND_URL"], "http://127.0.0.1:8101")
        self.assertEqual(probes[0]["start"], 8100)
        self.assertEqual(probes[1]["start"], 5200)
        self.assertEqual([item["host"] for item in probes], ["0.0.0.0", "0.0.0.0"])
        self.assertEqual(probes[1]["reserved"], {8101})

    def test_default_launch_remains_local(self):
        processes, probes = self.launch()
        backend_command, frontend_command = [item.args[0] for item in processes]
        self.assertEqual(self.option(backend_command, "--host"), "127.0.0.1")
        self.assertIn(self.option(frontend_command, "--host", "localhost"), {"localhost", "127.0.0.1"})
        self.assertIn("--reload", backend_command)
        self.assertEqual([item["start"] for item in probes], [8000, 5173])
        self.assertEqual([item["host"] for item in probes], ["127.0.0.1", "127.0.0.1"])
        self.assertEqual(processes[1].kwargs["env"]["VITE_BACKEND_URL"], "http://127.0.0.1:8000")

    def test_explicit_backend_host_takes_precedence_in_lan_mode(self):
        processes, probes = self.launch(
            ["--lan", "--host", "192.168.1.42"],
            {"BACKEND_HOST": "127.0.0.1"},
        )
        self.assertEqual(self.option(processes[0].args[0], "--host"), "192.168.1.42")
        self.assertEqual(self.option(processes[1].args[0], "--host"), "0.0.0.0")
        self.assertEqual(probes[0]["host"], "192.168.1.42")
        self.assertEqual(processes[1].kwargs["env"]["VITE_BACKEND_URL"], "http://192.168.1.42:8000")

    def test_environment_ports_and_backend_host_are_respected(self):
        processes, probes = self.launch(
            environment={"BACKEND_HOST": "192.168.1.42", "BACKEND_PORT": "8100", "FRONTEND_PORT": "5200"},
            selected_ports=(8100, 5200),
        )
        self.assertEqual([item["start"] for item in probes], [8100, 5200])
        self.assertEqual(probes[0]["host"], "192.168.1.42")
        self.assertEqual(self.option(processes[0].args[0], "--host"), "192.168.1.42")
        self.assertEqual(processes[1].kwargs["env"]["VITE_BACKEND_URL"], "http://192.168.1.42:8100")

    def test_cli_zero_port_is_rejected_instead_of_using_environment(self):
        for option in ("--backend-port", "--frontend-port"):
            with self.subTest(option=option), self.assertRaises(run.RunError):
                self.launch([option, "0"], {"BACKEND_PORT": "8100", "FRONTEND_PORT": "5200"})

    def test_invalid_port_values_are_rejected(self):
        for value in (-1, 0, 65536, "not-a-port"):
            with self.subTest(value=value), self.assertRaises(run.RunError):
                run.parse_port(value, "test port")

    def test_available_port_skips_conflicts_and_reserved_ports_on_requested_host(self):
        with patch.object(run.socket, "socket") as socket_factory:
            probe = socket_factory.return_value.__enter__.return_value
            probe.bind.side_effect = [OSError(errno.EADDRINUSE, "address in use"), None]
            self.assertEqual(run.available_port(8100, {8101}, host="0.0.0.0"), 8102)
            self.assertEqual(probe.bind.call_args_list, [call(("0.0.0.0", 8100)), call(("0.0.0.0", 8102))])

    def test_available_port_defaults_to_loopback(self):
        with patch.object(run.socket, "socket") as socket_factory:
            self.assertEqual(run.available_port(8100), 8100)
            socket_factory.return_value.__enter__.return_value.bind.assert_called_once_with(("127.0.0.1", 8100))

    def test_available_port_reports_exhaustion(self):
        with patch.object(run.socket, "socket") as socket_factory:
            socket_factory.return_value.__enter__.return_value.bind.side_effect = OSError(errno.EADDRINUSE, "address in use")
            with self.assertRaisesRegex(run.RunError, "No available TCP port"):
                run.available_port(65535, host="0.0.0.0")

    def test_invalid_bind_address_fails_without_scanning_all_ports(self):
        with patch.object(run.socket, "socket") as socket_factory:
            probe = socket_factory.return_value.__enter__.return_value
            probe.bind.side_effect = OSError(errno.EADDRNOTAVAIL, "address unavailable")
            with self.assertRaisesRegex(run.RunError, "Cannot bind to"):
                run.available_port(8100, host="192.168.1.42")
            probe.bind.assert_called_once_with(("192.168.1.42", 8100))

    def test_ipv6_bind_uses_ipv6_socket(self):
        with patch.object(run.socket, "socket") as socket_factory:
            self.assertEqual(run.available_port(8100, host="::"), 8100)
            socket_factory.assert_called_once_with(run.socket.AF_INET6, run.socket.SOCK_STREAM)

    def test_ipv6_proxy_url_brackets_backend_address(self):
        for host, expected in (("::", "[::1]"), ("::1", "[::1]")):
            with self.subTest(host=host):
                processes, _ = self.launch(["--host", host])
                self.assertEqual(processes[1].kwargs["env"]["VITE_BACKEND_URL"], f"http://{expected}:8000")


if __name__ == "__main__":
    unittest.main()
