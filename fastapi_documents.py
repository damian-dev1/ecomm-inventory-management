from __future__ import annotations

from datetime import datetime
from enum import Enum
from typing import Optional, List, Annotated
import json
from fastapi import FastAPI, HTTPException, Query, Request, Depends
from fastapi.responses import JSONResponse
from pydantic import BaseModel, Field, EmailStr
from sqlalchemy import func, text
from sqlmodel import SQLModel, Field as ORMField, Session, create_engine, select


# --------------------------------------------------------------------------------------
# DB Setup
# --------------------------------------------------------------------------------------

engine = create_engine("sqlite:///./orders.db", connect_args={"check_same_thread": False})

def get_session():
    with Session(engine) as session:
        yield session

# --------------------------------------------------------------------------------------
# ORM Models (no relationships; keep it dead simple)
# --------------------------------------------------------------------------------------

class OrderStatus(str, Enum):
    ORDER = "ORDER"
    ACKNOWLEDGED = "ACKNOWLEDGED"
    BACKORDERED = "BACKORDERED"
    DISPATCHED = "DISPATCHED"
    CANCELLATION_ACK = "CANCELLATION_ACK"

class Order(SQLModel, table=True):
    id: Optional[int] = ORMField(default=None, primary_key=True)
    url: Optional[str] = None
    retailer: Optional[str] = None
    order_reference: str = ORMField(index=True)
    order_date: datetime = ORMField(default_factory=datetime.utcnow)
    status: OrderStatus = ORMField(default=OrderStatus.ORDER, index=True)
    channel: Optional[str] = None
    purchase_order_reference: Optional[str] = None
    end_user_purchase_order_reference: Optional[str] = None
    additional_order_reference: Optional[str] = None
    comment: Optional[str] = ""
    test_flag: bool = False
    supplier: Optional[str] = None
    currency_code: Optional[str] = "AUD"
    subtotal: Optional[str] = None
    tax: Optional[str] = None
    total: Optional[str] = None

    # shipping address (flattened)
    ship_country: Optional[str] = None
    ship_line_1: Optional[str] = None
    ship_line_2: Optional[str] = None
    ship_city: Optional[str] = None
    ship_postal_code: Optional[str] = None
    ship_state: Optional[str] = None
    ship_phone: Optional[str] = None
    ship_full_name: Optional[str] = None
    ship_email: Optional[EmailStr] = None

    # retailer_data (minimal subset)
    retailer_name: Optional[str] = None
    delivery_service_code: Optional[str] = None

    # dispatch data
    dispatch_carrier: Optional[str] = None
    dispatch_tracking_number: Optional[str] = None
    dispatch_datetime: Optional[datetime] = None

class OrderItem(SQLModel, table=True):
    id: Optional[int] = ORMField(default=None, primary_key=True)
    order_id: int = ORMField(index=True, foreign_key="order.id")
    url: Optional[str] = None

    part_number: Optional[str] = None
    retailer_sku_reference: Optional[str] = None
    supplier_sku_reference: Optional[str] = None
    line_reference: Optional[str] = None
    quantity: int = 1
    name: Optional[str] = None
    description: Optional[str] = None
    status: Optional[str] = "ORDER"
    unit_cost_price: Optional[str] = None
    subtotal: Optional[str] = None
    tax: Optional[str] = None
    tax_rate: Optional[str] = None
    total: Optional[str] = None
    promised_date: Optional[datetime] = None
    retailer_additional_reference: Optional[str] = None

class Product(SQLModel, table=True):
    id: Optional[int] = ORMField(default=None, primary_key=True)
    sku: str = ORMField(index=True, unique=True)
    name: str
    stock: int = 0
    sap_article_id: Optional[str] = None  # NEW

class Supplier(SQLModel, table=True):
    id: Optional[int] = ORMField(default=None, primary_key=True)
    name: str = ORMField(index=True)
    uuid: Optional[str] = None
    postcode: Optional[str] = None
    country: Optional[str] = None
    categories_json: Optional[str] = None  # store list as JSON
    account_id: Optional[str] = None

# --------------------------------------------------------------------------------------
# Pydantic Schemas
# --------------------------------------------------------------------------------------

class Address(BaseModel):
    country: Optional[str] = None
    line_1: Optional[str] = None
    line_2: Optional[str] = None
    city: Optional[str] = None
    postal_code: Optional[str] = None
    state: Optional[str] = None
    phone: Optional[str] = None
    full_name: Optional[str] = None
    email: Optional[EmailStr] = None

class RetailerData(BaseModel):
    uuid: Optional[str] = None
    name: Optional[str] = None
    email: Optional[str] = None
    phone: Optional[str] = None
    tax_code: Optional[str] = None
    address: Optional[Address] = None

class OrderItemIn(BaseModel):
    url: Optional[str] = None
    part_number: Optional[str] = None
    retailer_sku_reference: Optional[str] = None
    supplier_sku_reference: Optional[str] = None
    line_reference: Optional[str] = None
    quantity: int = 1
    name: Optional[str] = None
    description: Optional[str] = None
    status: Optional[str] = "ORDER"
    unit_cost_price: Optional[str] = None
    subtotal: Optional[str] = None
    tax: Optional[str] = None
    tax_rate: Optional[str] = None
    total: Optional[str] = None
    promised_date: Optional[datetime] = None
    retailer_additional_reference: Optional[str] = None

class OrderIn(BaseModel):
    order_reference: str
    order_date: Optional[datetime] = None
    status: Optional[OrderStatus] = OrderStatus.ORDER
    comment: Optional[str] = ""
    test_flag: bool = False
    currency_code: Optional[str] = "AUD"
    subtotal: Optional[str] = None
    tax: Optional[str] = None
    total: Optional[str] = None
    shipping_address: Optional[Address] = None
    retailer_data: Optional[RetailerData] = None
    supplier: Optional[str] = None
    retailer: Optional[str] = None
    items: List[OrderItemIn] = Field(default_factory=list)

class OrderItemOut(OrderItemIn):
    url: Optional[str] = None

class OrderOut(BaseModel):
    url: Optional[str] = None
    retailer: Optional[str] = None
    order_reference: str
    order_date: datetime
    status: OrderStatus
    channel: Optional[str] = None
    purchase_order_reference: Optional[str] = None
    end_user_purchase_order_reference: Optional[str] = None
    additional_order_reference: Optional[str] = None
    comment: Optional[str] = ""
    test_flag: bool = False
    supplier: Optional[str] = None
    items: List[OrderItemOut] = Field(default_factory=list)
    currency_code: Optional[str] = "AUD"
    subtotal: Optional[str] = None
    tax: Optional[str] = None
    total: Optional[str] = None
    shipping_address: Optional[Address] = None
    retailer_data: Optional[RetailerData] = None
    delivery_service_code: Optional[str] = None

class PageEnvelope(BaseModel):
    count: int
    next: Optional[str]
    previous: Optional[str]
    results: List[OrderOut]

class CarrierOut(BaseModel):
    code: str
    name: str
    services: List[str] = Field(default_factory=list)

class ProductIn(BaseModel):
    sku: str
    name: str
    stock: int = 0
    sap_article_id: Optional[str] = None  # NEW (optional on create)

class ProductOut(ProductIn):
    pass


class StockPatch(BaseModel):
    stock: int

class BackorderIn(BaseModel):
    reason: Optional[str] = None

class DispatchIn(BaseModel):
    carrier: str
    tracking_number: Optional[str] = None
    service: Optional[str] = None

class SupplierOut(BaseModel):
    url: str
    name: str
    uuid: Optional[str] = None
    postcode: Optional[str] = None
    country: Optional[str] = None
    categories: List[str] = Field(default_factory=list)
    account_id: Optional[str] = None

class SupplierEnvelope(BaseModel):
    count: int
    next: Optional[str]
    previous: Optional[str]
    results: List[SupplierOut]

# --------------------------------------------------------------------------------------
# App
# --------------------------------------------------------------------------------------

app = FastAPI(title="Order & Product API", version="1.0.2")

@app.on_event("startup")
def on_startup():
    SQLModel.metadata.create_all(engine)

    # --- tiny migration for products.sap_article_id (keep if you already added it) ---
    with engine.connect() as conn:
        cols = [r[1] for r in conn.exec_driver_sql("PRAGMA table_info('product')").fetchall()]
        if "sap_article_id" not in cols:
            conn.execute(text("ALTER TABLE product ADD COLUMN sap_article_id VARCHAR"))
            conn.commit()

    # --- optional seed for suppliers if table empty ---
    with Session(engine) as s:
        any_supplier = s.exec(select(func.count()).select_from(Supplier)).one()
        if isinstance(any_supplier, tuple):
            any_supplier = any_supplier[0]
        if int(any_supplier) == 0:
            s.add(
                Supplier(
                    name="test-prod-restapi-supplier",
                    uuid="9442a2e1-6ed1-405c-b8b8-190f13b4cc70",
                    postcode="RG1 3AR",
                    country="GB",
                    categories_json=json.dumps([]),
                    account_id="24680",
                )
            )
            s.commit()



# --------------------------------------------------------------------------------------
# Utilities
# --------------------------------------------------------------------------------------

def _model_dump(m):
    return m.model_dump() if hasattr(m, "model_dump") else m.dict()

def _items_for(session: Session, order_id: int) -> List[OrderItem]:
    return session.exec(select(OrderItem).where(OrderItem.order_id == order_id)).all()

def _supplier_to_out(s: Supplier, base: str, namespace: str) -> SupplierOut:
    # namespace: "api" (products) or "restapi" (orders)
    return SupplierOut(
        url=f"{base}/{namespace}/v4/suppliers/{s.id}/",
        name=s.name,
        uuid=s.uuid,
        postcode=s.postcode,
        country=s.country,
        categories=json.loads(s.categories_json) if s.categories_json else [],
        account_id=s.account_id,
    )

def _supplier_page_links(request: Request, count: int, limit: int, offset: int) -> tuple[Optional[str], Optional[str]]:
    def build(off: int) -> str:
        q = dict(request.query_params)
        q["limit"] = str(limit)
        q["offset"] = str(off)
        base = str(request.url).split("?")[0]
        from urllib.parse import urlencode
        return f"{base}?{urlencode(q)}"
    next_url = build(offset + limit) if offset + limit < count else None
    prev_url = build(max(offset - limit, 0)) if offset > 0 else None
    return next_url, prev_url

def to_order_out(order: Order, items: List[OrderItem]) -> OrderOut:
    shipping = Address(
        country=order.ship_country, line_1=order.ship_line_1, line_2=order.ship_line_2,
        city=order.ship_city, postal_code=order.ship_postal_code, state=order.ship_state,
        phone=order.ship_phone, full_name=order.ship_full_name, email=order.ship_email,
    )
    retailer_data = RetailerData(name=order.retailer_name) if order.retailer_name else None
    return OrderOut(
        url=order.url,
        retailer=order.retailer,
        order_reference=order.order_reference,
        order_date=order.order_date,
        status=order.status,
        comment=order.comment or "",
        test_flag=order.test_flag,
        supplier=order.supplier,
        items=[OrderItemOut(**(_model_dump(i) | {"url": i.url})) for i in items],
        currency_code=order.currency_code,
        subtotal=order.subtotal,
        tax=order.tax,
        total=order.total,
        shipping_address=shipping if any(_model_dump(shipping).values()) else None,
        retailer_data=retailer_data,
        delivery_service_code=order.delivery_service_code,
    )

def page_links(request: Request, count: int, limit: int, offset: int) -> tuple[Optional[str], Optional[str]]:
    def build(off: int) -> str:
        q = dict(request.query_params)
        q["limit"] = str(limit)
        q["offset"] = str(off)
        base = str(request.url).split("?")[0]
        from urllib.parse import urlencode
        return f"{base}?{urlencode(q)}"
    next_url = build(offset + limit) if offset + limit < count else None
    prev_url = build(max(offset - limit, 0)) if offset > 0 else None
    return next_url, prev_url

# --------------------------------------------------------------------------------------
# Orders
# --------------------------------------------------------------------------------------

@app.post("/orders", response_model=OrderOut, status_code=201, summary="Create order")
def create_order(payload: OrderIn, session: Session = Depends(get_session), request: Request = None):
    exists = session.exec(select(Order).where(Order.order_reference == payload.order_reference)).first()
    if exists:
        raise HTTPException(status_code=409, detail="order_reference already exists")

    o = Order(
        order_reference=payload.order_reference,
        order_date=payload.order_date or datetime.utcnow(),
        status=payload.status or OrderStatus.ORDER,
        comment=payload.comment or "",
        test_flag=payload.test_flag,
        currency_code=payload.currency_code,
        subtotal=payload.subtotal,
        tax=payload.tax,
        total=payload.total,
        supplier=payload.supplier,
        retailer=payload.retailer,
    )
    if payload.shipping_address:
        sa = payload.shipping_address
        o.ship_country = sa.country
        o.ship_line_1 = sa.line_1
        o.ship_line_2 = sa.line_2
        o.ship_city = sa.city
        o.ship_postal_code = sa.postal_code
        o.ship_state = sa.state
        o.ship_phone = sa.phone
        o.ship_full_name = sa.full_name
        o.ship_email = sa.email
    if payload.retailer_data:
        o.retailer_name = payload.retailer_data.name

    session.add(o)
    session.commit()
    session.refresh(o)

    # Create items
    items: List[OrderItem] = []
    for it in payload.items:
        d = _model_dump(it)
        item = OrderItem(order_id=o.id, **d)
        session.add(item)
        items.append(item)
    session.commit()
    for it in items:
        session.refresh(it)

    # canonical URLs
    base = str(request.base_url).rstrip("/")
    o.url = f"{base}/orders/{o.id}"
    session.add(o)
    for it in items:
        it.url = f"{base}/orders/{o.id}/items/{it.id}"
        session.add(it)
    session.commit()

    return to_order_out(o, items)

@app.get("/orders", response_model=PageEnvelope, summary="List Orders")
def list_orders(
    request: Request,
    session: Session = Depends(get_session),
    limit: Annotated[int, Query(ge=1, le=200)] = 50,
    offset: Annotated[int, Query(ge=0)] = 0,
    status: Optional[OrderStatus] = Query(None),
):
    stmt = select(Order)
    if status:
        stmt = stmt.where(Order.status == status)

    total = session.exec(select(func.count()).select_from(Order)).one()
    if isinstance(total, tuple):  # safety for driver/dialect differences
        total = total[0]

    rows = session.exec(stmt.order_by(Order.id).limit(limit).offset(offset)).all()

    results: List[OrderOut] = []
    for o in rows:
        items = _items_for(session, o.id)
        results.append(to_order_out(o, items))

    next_url, prev_url = page_links(request, total, limit, offset)
    return PageEnvelope(count=int(total), next=next_url, previous=prev_url, results=results)

@app.get("/orders/{order_id}", response_model=OrderOut, summary="View Order details")
def get_order(order_id: int, session: Session = Depends(get_session), request: Request = None):
    o = session.get(Order, order_id)
    if not o:
        raise HTTPException(status_code=404, detail="Order not found")
    items = _items_for(session, order_id)
    return to_order_out(o, items)

@app.post("/orders/{order_id}/acknowledge", response_model=OrderOut, summary="Acknowledge an order")
def acknowledge_order(order_id: int, session: Session = Depends(get_session)):
    o = session.get(Order, order_id)
    if not o:
        raise HTTPException(status_code=404, detail="Order not found")
    if o.status not in (OrderStatus.ORDER, OrderStatus.BACKORDERED):
        raise HTTPException(status_code=409, detail=f"Cannot acknowledge from status {o.status}")
    o.status = OrderStatus.ACKNOWLEDGED
    session.add(o)
    session.commit()
    return to_order_out(o, _items_for(session, order_id))

@app.post("/orders/{order_id}/backorder", response_model=OrderOut, summary="Backorder an order")
def backorder_order(order_id: int, payload: BackorderIn, session: Session = Depends(get_session)):
    o = session.get(Order, order_id)
    if not o:
        raise HTTPException(status_code=404, detail="Order not found")
    if o.status in (OrderStatus.DISPATCHED, OrderStatus.CANCELLATION_ACK):
        raise HTTPException(status_code=409, detail=f"Cannot backorder from status {o.status}")
    o.status = OrderStatus.BACKORDERED
    o.comment = (o.comment or "")
    if payload.reason:
        o.comment += f" | Backorder: {payload.reason}"
    session.add(o)
    session.commit()
    return to_order_out(o, _items_for(session, order_id))

CARRIERS = [
    {"code": "AUSPOST_PARCEL", "name": "Australia Post Parcels", "services": ["STANDARD", "EXPRESS"]},
    {"code": "STARTRACK", "name": "StarTrack", "services": ["PRIORITY", "AUTHORITY_TO_LEAVE"]},
    {"code": "TOLL", "name": "Toll IPEC", "services": ["ROAD", "PRIORITY"]},
]

@app.get("/carriers", response_model=List[CarrierOut], summary="List carriers")
def list_carriers():
    return CARRIERS

@app.post("/orders/{order_id}/dispatch", response_model=OrderOut, summary="Dispatch an order")
def dispatch_order(order_id: int, payload: DispatchIn, session: Session = Depends(get_session)):
    o = session.get(Order, order_id)
    if not o:
        raise HTTPException(status_code=404, detail="Order not found")
    if o.status not in (OrderStatus.ORDER, OrderStatus.ACKNOWLEDGED, OrderStatus.BACKORDERED):
        raise HTTPException(status_code=409, detail=f"Cannot dispatch from status {o.status}")
    if payload.carrier not in {c["code"] for c in CARRIERS}:
        raise HTTPException(status_code=400, detail="Unknown carrier code")
    o.status = OrderStatus.DISPATCHED
    o.delivery_service_code = payload.service
    o.dispatch_carrier = payload.carrier
    o.dispatch_tracking_number = payload.tracking_number
    o.dispatch_datetime = datetime.utcnow()
    session.add(o)
    session.commit()
    return to_order_out(o, _items_for(session, order_id))

@app.post("/orders/{order_id}/cancellation/ack", response_model=OrderOut, summary="Acknowledge a cancellation")
def acknowledge_cancellation(order_id: int, session: Session = Depends(get_session)):
    o = session.get(Order, order_id)
    if not o:
        raise HTTPException(status_code=404, detail="Order not found")
    if o.status == OrderStatus.DISPATCHED:
        raise HTTPException(status_code=409, detail="Cannot acknowledge cancellation after dispatch")
    o.status = OrderStatus.CANCELLATION_ACK
    session.add(o)
    session.commit()
    return to_order_out(o, _items_for(session, order_id))

# --------------------------------------------------------------------------------------
# Products
# --------------------------------------------------------------------------------------

# @app.post("/products", response_model=ProductOut, status_code=201, summary="Create product")
# def create_product(payload: ProductIn, session: Session = Depends(get_session)):
#     exists = session.exec(select(Product).where(Product.sku == payload.sku)).first()
#     if exists:
#         raise HTTPException(status_code=409, detail="SKU already exists")
#     p = Product(sku=payload.sku, name=payload.name, stock=payload.stock)
#     session.add(p)
#     session.commit()
#     session.refresh(p)
#     return ProductOut(sku=p.sku, name=p.name, stock=p.stock)

@app.post("/products", response_model=ProductOut, status_code=201, summary="Create product")
def create_product(payload: ProductIn, session: Session = Depends(get_session)):
    exists = session.exec(select(Product).where(Product.sku == payload.sku)).first()
    if exists:
        raise HTTPException(status_code=409, detail="SKU already exists")
    p = Product(
        sku=payload.sku,
        name=payload.name,
        stock=payload.stock,
        sap_article_id=payload.sap_article_id,   # NEW
    )
    session.add(p)
    session.commit()
    session.refresh(p)
    return ProductOut(**_model_dump(p))


# @app.get("/products", response_model=List[ProductOut], summary="List products")
# def list_products(sku: Optional[str] = None, session: Session = Depends(get_session)):
#     stmt = select(Product)
#     if sku:
#         stmt = stmt.where(Product.sku == sku)
#     rows = session.exec(stmt.order_by(Product.sku)).all()
#     return [ProductOut(sku=r.sku, name=r.name, stock=r.stock) for r in rows]


@app.get("/products", response_model=List[ProductOut], summary="List products")
def list_products(sku: Optional[str] = None, session: Session = Depends(get_session)):
    stmt = select(Product)
    if sku:
        stmt = stmt.where(Product.sku == sku)
    rows = session.exec(stmt.order_by(Product.sku)).all()
    return [ProductOut(**_model_dump(r)) for r in rows]   # keep extras (sap_article_id)


class SapPatch(BaseModel):
    sap_article_id: str | int

@app.patch("/products/{sku}/sap", response_model=ProductOut, summary="Update SAP Article ID")
def update_sap_id(sku: str, payload: SapPatch, session: Session = Depends(get_session)):
    p = session.exec(select(Product).where(Product.sku == sku)).first()
    if not p:
        raise HTTPException(status_code=404, detail="Product not found")

    # Coerce to string to match VS behavior (string in response)
    p.sap_article_id = str(payload.sap_article_id).strip()
    if not p.sap_article_id:
        raise HTTPException(status_code=400, detail="sap_article_id cannot be empty")

    session.add(p)
    session.commit()
    session.refresh(p)
    return ProductOut(**_model_dump(p))


@app.patch("/products/{sku}/stock", response_model=ProductOut, summary="Update stock")
def update_stock(sku: str, payload: StockPatch, session: Session = Depends(get_session)):
    p = session.exec(select(Product).where(Product.sku == sku)).first()
    if not p:
        raise HTTPException(status_code=404, detail="Product not found")
    if payload.stock < 0:
        raise HTTPException(status_code=400, detail="Stock cannot be negative")
    p.stock = payload.stock
    session.add(p)
    session.commit()
    session.refresh(p)
    return ProductOut(sku=p.sku, name=p.name, stock=p.stock)

@app.get("/api/v4/suppliers/", response_model=SupplierEnvelope, summary="List suppliers (products namespace)")
def list_suppliers_products(
    request: Request,
    session: Session = Depends(get_session),
    limit: Annotated[int, Query(ge=1, le=2000)] = 1000,
    offset: Annotated[int, Query(ge=0)] = 0,
    name: Optional[str] = Query(None, description="Optional case-insensitive name filter")):
    stmt = select(Supplier)
    if name:
        # naive contains filter (SQLite: LIKE is case-insensitive by default)
        stmt = stmt.where(Supplier.name.like(f"%{name}%"))

    total = session.exec(select(func.count()).select_from(Supplier)).one()
    if isinstance(total, tuple):
        total = total[0]

    rows = session.exec(stmt.order_by(Supplier.id).limit(limit).offset(offset)).all()

    base = str(request.base_url).rstrip("/")
    results = [_supplier_to_out(s, base, "api") for s in rows]
    next_url, prev_url = _supplier_page_links(request, int(total), limit, offset)
    return SupplierEnvelope(count=int(total), next=next_url, previous=prev_url, results=results)

# --------------------------------------------------------------------------------------
# Error handler
# --------------------------------------------------------------------------------------

@app.exception_handler(HTTPException)
def http_exc_handler(_, exc: HTTPException):
    return JSONResponse(status_code=exc.status_code, content={"detail": exc.detail})
