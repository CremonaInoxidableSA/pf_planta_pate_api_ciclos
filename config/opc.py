from opcua import Client
from opcua.common.subscription import SubHandler
import logging
import asyncio

logger = logging.getLogger("uvicorn")


class OPCUASubscriptionHandler(SubHandler):
    """
    Recibe notificaciones push del servidor OPC UA.
    Cada vez que un nodo suscripto cambia, datachange_notification
    actualiza el cache en memoria \u2014 sin ning�n round-trip de red.

    El cache se indexa por node_id (string): node.nodeid.to_string()
    """

    def __init__(self):
        self._cache: dict[str, object] = {}

    def get_value(self, node_id: str):
        return self._cache.get(node_id)

    def get_all(self) -> dict:
        return dict(self._cache)

    # \u2705 NUEVO: limpia el cache para evitar servir valores obsoletos
    def clear(self):
        self._cache.clear()
        logger.info("Cache del handler OPC limpiado.")

    # --- Llamados autom�ticamente por la librer�a opcua ---

    def datachange_notification(self, node, val, data):
        node_id = node.nodeid.to_string()
        self._cache[node_id] = val

    def event_notification(self, event):
        logger.debug(f"Evento OPC recibido: {event}")

    def status_change_notification(self, status):
        logger.warning(f"Cambio de estado en suscripci�n OPC: {status}")


class OPCUAClient:
    def __init__(self, server_url, max_retries=3, retry_delay=3):
        self.server_url   = server_url
        self.client       = None
        self.max_retries  = max_retries
        self.retry_delay  = retry_delay
        self.connected    = False

        self._subscription    = None
        self.handler          = OPCUASubscriptionHandler()
        self._nodos_suscritos: list = []   # nodos Node ya navegados

    # ------------------------------------------------------------------
    # Conexi�n / desconexi�n
    # ------------------------------------------------------------------

    async def connect(self):
        """
        Intenta conectarse al servidor OPC.

        Devuelve:
            True: conexi�n validada correctamente.
            False: se agotaron los intentos.
        """

        # Si ya existe una conexi�n realmente funcional, no crear otra.
        if await self.ping():
            return True

        # Limpiar cualquier cliente anterior que haya quedado incompleto.
        if self.client is not None:
            await self.disconnect()

        self.connected = False
        self.client = None

        for intento in range(1, self.max_retries + 1):
            candidate = Client(self.server_url)

            try:
                def _conectar_y_validar():
                    candidate.connect()

                    # Validar la conexi�n mediante una lectura OPC real.
                    candidate.get_root_node().get_browse_name()

                await asyncio.to_thread(_conectar_y_validar)

                # Publicar el cliente solamente despu�s de validarlo.
                self.client = candidate
                self.connected = True

                logger.info("\u2705 Conectado al servidor OPC UA.")
                return True

            except Exception as e:
                self.connected = False
                self.client = None

                # El cliente pudo quedar parcialmente conectado.
                try:
                    await asyncio.to_thread(candidate.disconnect)
                except Exception:
                    pass

                logger.error(
                    "\U0001f504 Error al conectar OPC UA. Intento %s/%s: %s",
                    intento,
                    self.max_retries,
                    e,
                )

                if intento < self.max_retries:
                    await asyncio.sleep(self.retry_delay)

        logger.warning(
            "\u26a0\ufe0f No se pudo conectar al servidor despu�s de varios intentos."
        )
        return False

    async def disconnect(self):
        """
        Cancela la suscripci�n, limpia el cach� y cierra la conexi�n.
        """

        # Guardar la referencia local antes de limpiar el estado p�blico.
        client_actual = self.client

        # Marcar inmediatamente el cliente como no disponible.
        self.connected = False
        self.client = None

        await self._cancelar_suscripcion()

        # Evitar reutilizar valores anteriores despu�s de reconectar.
        self.handler.clear()
        self._nodos_suscritos = []

        if client_actual:
            try:
                await asyncio.to_thread(client_actual.disconnect)
            except Exception:
                pass

            logger.warning("\u26a0\ufe0f Conexi�n OPC UA cerrada.")

    async def reconnect(self):
        """
        Desconecta y vuelve a conectar.

        La suscripci�n no se reconstruye aqu� porque los Node anteriores
        dejan de ser v�lidos despu�s de una reconexi�n.
        """

        logger.info("\U0001f504 Intentando reconectar al servidor OPC UA...")

        await self.disconnect()
        return await self.connect()

    # ------------------------------------------------------------------
    # Health check \u2014 para el monitor en main.py
    # ------------------------------------------------------------------

    async def ping(self) -> bool:
        """
        Verifica que la conexi�n OPC UA siga activa haciendo una lectura
        m�nima (browse_name del root node). Devuelve True si la conexi�n
        est� viva, False si cay�.
        """
        if not self.client or not self.connected:
            return False
        try:
            def _check():
                self.client.get_root_node().get_browse_name()
            await asyncio.to_thread(_check)
            return True
        except Exception as e:
            logger.warning(f"\u26a0\ufe0f Ping OPC fallido: {e}")
            return False

    # ------------------------------------------------------------------
    # Suscripci�n \u2014 recibe nodos YA NAVEGADOS (objetos Node de opcua)
    # ------------------------------------------------------------------

    async def suscribir_nodos(self, nodos: list, period_ms: int = 500):
        """
        Crea una suscripci�n utilizando nodos reci�n navegados.
        """

        if not self.connected or not self.client:
            logger.error(
                "No se puede suscribir: cliente OPC no conectado."
            )
            return False

        if not nodos:
            logger.error(
                "No se puede suscribir: la lista de nodos est� vac�a."
            )
            return False

        await self._cancelar_suscripcion()

        cliente_actual = self.client

        def _crear():
            sub = cliente_actual.create_subscription(
                period_ms,
                self.handler,
            )

            try:
                sub.subscribe_data_change(nodos)
            except Exception:
                try:
                    sub.delete()
                except Exception:
                    pass
                raise

            return sub

        try:
            nueva_suscripcion = await asyncio.to_thread(_crear)

            self._subscription = nueva_suscripcion
            self._nodos_suscritos = list(nodos)

            logger.info(
                "\u2705 Suscripci�n OPC activa: %s nodos, per�odo %s ms",
                len(nodos),
                period_ms,
            )
            return True

        except Exception as e:
            self._subscription = None
            self._nodos_suscritos = []

            logger.error(
                "Error al crear suscripci�n OPC UA: %s",
                e,
            )
            return False

    async def _cancelar_suscripcion(self):
        if self._subscription:
            try:
                await asyncio.to_thread(self._subscription.delete)
            except Exception:
                pass
            self._subscription = None

    # ------------------------------------------------------------------
    # Lectura puntual (se mantiene para recetario y compatibilidad)
    # ------------------------------------------------------------------

    def read_node(self, node_id: str):
        if not self.client or not self.connected:
            raise Exception("\u26a0\ufe0f Cliente OPC UA no conectado.")
        return self.client.get_node(node_id).get_value()

    def get_objects_node(self):
        if not self.client or not self.connected:
            raise Exception("\u26a0\ufe0f Cliente OPC UA no conectado.")
        return self.client.get_objects_node()

    async def get_objects_nodos(self):
        if not self.client or not self.connected:
            raise Exception("\u26a0\ufe0f Cliente OPC UA no conectado.")
        return self.client.get_root_node()

    async def handle_reconnect(self):
        """
        Intenta recuperar solamente la conexi�n.

        La navegaci�n y suscripci�n ser�n administradas por main.py.
        """

        try:
            return await self.reconnect()

        except Exception as e:
            logger.error(
                "\u26a0\ufe0f Error al intentar reconectar: %s",
                e,
            )
            return False