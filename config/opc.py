from opcua import Client
from opcua.common.subscription import SubHandler
import logging
import asyncio

logger = logging.getLogger("uvicorn")


class OPCUASubscriptionHandler(SubHandler):
    """
    Recibe notificaciones push del servidor OPC UA.
    Cada vez que un nodo suscripto cambia, datachange_notification
    actualiza el cache en memoria — sin ningún round-trip de red.

    El cache se indexa por node_id (string): node.nodeid.to_string()
    """

    def __init__(self):
        self._cache: dict[str, object] = {}

    def get_value(self, node_id: str):
        return self._cache.get(node_id)

    def get_all(self) -> dict:
        return dict(self._cache)

    # ✅ NUEVO: limpia el cache para evitar servir valores obsoletos
    def clear(self):
        self._cache.clear()
        logger.info("Cache del handler OPC limpiado.")

    # --- Llamados automáticamente por la librería opcua ---

    def datachange_notification(self, node, val, data):
        node_id = node.nodeid.to_string()
        self._cache[node_id] = val

    def event_notification(self, event):
        logger.debug(f"Evento OPC recibido: {event}")

    def status_change_notification(self, status):
        logger.warning(f"Cambio de estado en suscripción OPC: {status}")


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
    # Conexión / desconexión
    # ------------------------------------------------------------------

    async def connect(self):
        retries = 0
        while retries < self.max_retries:
            try:
                if self.client is None or not self.connected:
                    self.client = Client(self.server_url)
                    await asyncio.to_thread(self.client.connect)
                    self.connected = True
                    logger.info("✅ Conectado al servidor OPC UA.")
                    return
            except Exception as e:
                retries += 1
                logger.error(
                    f"🔄 Error al conectar OPC UA. Intento {retries}/{self.max_retries}: {e}"
                )
                await asyncio.sleep(self.retry_delay)
        logger.warning("⚠️ No se pudo conectar al servidor después de varios intentos.")

    async def disconnect(self):
        """
        Cancela la suscripción, limpia el cache y cierra la conexión TCP.
        Limpiar el cache aquí es crítico: evita que datosGenerales()
        siga sirviendo valores obsoletos de la conexión anterior.
        """
        await self._cancelar_suscripcion()

        # ✅ NUEVO: limpiar cache al desconectar
        self.handler.clear()
        self._nodos_suscritos = []

        if self.client and self.connected:
            try:
                await asyncio.to_thread(self.client.disconnect)
            except Exception:
                pass
            logger.warning("⚠️ Conexión OPC UA cerrada.")
            self.connected = False
            self.client    = None

    async def reconnect(self):
        """
        Solo desconecta y vuelve a conectar.
        NO re-suscribe: esa responsabilidad queda en la capa de servicio
        (ObtenerNodosOpcUA.iniciar_suscripcion), que debe navegar el árbol
        de nuevo con nodos frescos de la nueva conexión.
        """
        logger.info("🔄 Intentando reconectar al servidor OPC UA...")
        await self.disconnect()
        await self.connect()
        # ✅ CORRECCIÓN: NO llamar a suscribir_nodos aquí.
        # Los Node del árbol anterior son inválidos en la nueva conexión.
        # main.py llamará a dGeneral.iniciar_suscripcion() después de reconnect().

    # ------------------------------------------------------------------
    # Health check — para el monitor en main.py
    # ------------------------------------------------------------------

    async def ping(self) -> bool:
        """
        Verifica que la conexión OPC UA siga activa haciendo una lectura
        mínima (browse_name del root node). Devuelve True si la conexión
        está viva, False si cayó.
        """
        if not self.client or not self.connected:
            return False
        try:
            def _check():
                self.client.get_root_node().get_browse_name()
            await asyncio.to_thread(_check)
            return True
        except Exception as e:
            logger.warning(f"⚠️ Ping OPC fallido: {e}")
            return False

    # ------------------------------------------------------------------
    # Suscripción — recibe nodos YA NAVEGADOS (objetos Node de opcua)
    # ------------------------------------------------------------------

    async def suscribir_nodos(self, nodos: list, period_ms: int = 500):
        """
        Crea (o re-crea) la suscripción OPC UA.

        Args:
            nodos:     Lista de objetos Node ya navegados desde el árbol OPC.
            period_ms: Intervalo de publicación en ms (default 500).
        """
        if not self.connected or not self.client:
            logger.error("No se puede suscribir: cliente OPC no conectado.")
            return

        await self._cancelar_suscripcion()

        def _crear():
            sub = self.client.create_subscription(period_ms, self.handler)
            sub.subscribe_data_change(nodos)
            return sub

        try:
            self._subscription    = await asyncio.to_thread(_crear)
            self._nodos_suscritos = list(nodos)
            logger.info(
                f"✅ Suscripción OPC activa: {len(nodos)} nodos, período {period_ms} ms"
            )
        except Exception as e:
            logger.error(f"Error al crear suscripción OPC UA: {e}")

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
            raise Exception("⚠️ Cliente OPC UA no conectado.")
        return self.client.get_node(node_id).get_value()

    def get_objects_node(self):
        if not self.client or not self.connected:
            raise Exception("⚠️ Cliente OPC UA no conectado.")
        return self.client.get_objects_node()

    async def get_objects_nodos(self):
        if not self.client or not self.connected:
            raise Exception("⚠️ Cliente OPC UA no conectado.")
        return self.client.get_root_node()

    async def handle_reconnect(self):
        """
        Punto de entrada de reconexión de emergencia desde datosGenerales().
        Solo intenta reconectar el transporte; la re-suscripción completa
        queda en manos del monitor_opc() de main.py.
        """
        try:
            await self.reconnect()
        except Exception as e:
            logger.error(f"⚠️ Error al intentar reconectar: {e}")